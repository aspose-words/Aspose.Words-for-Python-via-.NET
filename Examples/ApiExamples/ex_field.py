import io
import unittest

import api_example_base as aeb

import aspose.words as aw
from document_helper import DocumentHelper


class ExField(aeb.ApiExampleBase):

    def test_get_field_from_document(self):
        # ExStart
        # ExFor:FieldType
        # ExFor:FieldChar
        # ExFor:FieldChar.field_type
        # ExFor:FieldChar.is_dirty
        # ExFor:FieldChar.is_locked
        # ExFor:FieldChar.get_field
        # ExFor:Field.is_locked
        # ExSummary:Shows how to work with a FieldStart node.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        field = builder.insert_field(aw.fields.FieldType.FIELD_DATE, True)
        field.format.date_time_format = "dddd, MMMM dd, yyyy"
        field.update()

        field_start = field.start

        self.assertEqual(aw.fields.FieldType.FIELD_DATE, field_start.field_type)
        self.assertEqual(False, field_start.is_dirty)
        self.assertEqual(False, field_start.is_locked)

        # Retrieve the facade object which represents the field in the document.
        field = field_start.get_field()

        self.assertEqual(False, field.is_locked)
        self.assertEqual(" DATE  \\@ \"dddd, MMMM dd, yyyy\"", field.get_field_code())

        # Update the field to show the current date.
        field.update()
        # ExEnd

        doc = DocumentHelper.save_open(doc)

        # TestUtil.verify_field(FieldType.field_date, " DATE  \\@ \"dddd, MMMM dd, yyyy\"", DateTime.now.to_string("dddd, MMMM dd, yyyy"), doc.range.fields[0])

    def test_get_field_code(self):
        # ExStart
        # ExFor:Field.get_field_code
        # ExFor:Field.get_field_code(bool)
        # ExSummary:Shows how to get a field's field code.
        # Open a document which contains a MERGEFIELD inside an IF field.
        doc = aw.Document(aeb.my_dir + "Nested fields.docx")
        field_if = doc.range.fields[0]

        # There are two ways of getting a field's field code:
        # 1 -  Omit its inner fields:
        self.assertEqual(" IF  > 0 \" (surplus of ) \" \"\" ", field_if.get_field_code(False))

        # 2 -  Include its inner fields:
        self.assertEqual(
            " IF \u0013 MERGEFIELD NetIncome \u0014\u0015 > 0 \" (surplus of \u0013 MERGEFIELD  NetIncome \\f $ \u0014\u0015) \" \"\" ",
            field_if.get_field_code(True))

        # By default, the GetFieldCode method displays inner fields.
        self.assertEqual(field_if.get_field_code(), field_if.get_field_code(True))

        # ExEnd

    def test_display_result(self):
        # ExStart
        # ExFor:Field.display_result
        # ExSummary:Shows how to get the real text that a field displays in the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("This document was written by ")
        field_author = builder.insert_field(aw.fields.FieldType.FIELD_AUTHOR, True).as_field_author()
        field_author.author_name = "John Doe"

        # We can use the DisplayResult property to verify what exact text
        # a field would display in its place in the document.
        self.assertEqual("", field_author.display_result)

        # Fields do not maintain accurate result values in real-time.
        # To make sure our fields display accurate results at any given time,
        # such as right before a save operation, we need to update them manually.
        field_author.update()

        self.assertEqual("John Doe", field_author.display_result)

        doc.save(aeb.artifacts_dir + "Field.display_result.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "Field.display_result.docx")

        self.assertEqual("John Doe", doc.range.fields[0].display_result)

    def test_create_with_field_builder(self):
        # ExStart
        # ExFor:FieldBuilder.#ctor(FieldType)
        # ExFor:FieldBuilder.build_and_insert(Inline)
        # ExSummary:Shows how to create and insert a field using a field builder.
        doc = aw.Document()

        # A convenient way of adding text content to a document is with a document builder.
        builder = aw.DocumentBuilder(doc)
        builder.write(" Hello world! This text is one Run, which is an inline node.")

        # Fields have their builder, which we can use to construct a field code piece by piece.
        # In this case, we will construct a BARCODE field representing a US postal code,
        # and then insert it in front of a Run.
        field_builder = aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_BARCODE)
        field_builder.add_argument("90210")
        field_builder.add_switch("\\f", "A")
        field_builder.add_switch("\\u")

        field_builder.build_and_insert(doc.first_section.body.first_paragraph.runs[0])

        doc.update_fields()
        doc.save(aeb.artifacts_dir + "Field.create_with_field_builder.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "Field.create_with_field_builder.docx")

        # TestUtil.verify_field(FieldType.field_barcode, " BARCODE 90210 \\f A \\u ", string.empty, doc.range.fields[0])

        self.assertEqual(doc.first_section.body.first_paragraph.runs[11].previous_sibling, doc.range.fields[0].end)
        self.assertEqual(
            "ControlChar.field_start_char BARCODE 90210 \\f A \\u ControlChar.field_end_char Hello world! This text is one Run, which is an inline node.",
            doc.get_text().strip())

    def test_rev_num(self):
        # ExStart
        # ExFor:BuiltInDocumentProperties.revision_number
        # ExFor:FieldRevNum
        # ExSummary:Shows how to work with REVNUM fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Current revision #")

        # Insert a REVNUM field, which displays the document's current revision number property.
        field = builder.insert_field(aw.fields.FieldType.FIELD_REVISION_NUM, True).as_field_rev_num()

        self.assertEqual(" REVNUM ", field.get_field_code())
        self.assertEqual("1", field.result)
        self.assertEqual(1, doc.built_in_document_properties.revision_number)

        # This property counts how many times a document has been saved in Microsoft Word,
        # and is unrelated to tracked revisions. We can find it by right clicking the document in Windows Explorer
        # via Properties -> Details. We can update this property manually.
        doc.built_in_document_properties.revision_number += 1
        self.assertEqual("1", field.result)  # ExSkip
        field.update()

        self.assertEqual("2", field.result)
        # ExEnd

        doc = DocumentHelper.save_open(doc)
        self.assertEqual(2, doc.built_in_document_properties.revision_number)

        # TestUtil.verify_field(FieldType.field_revision_num, " REVNUM ", "2", doc.range.fields[0])

    def test_insert_field_none(self):
        # ExStart
        # ExFor:FieldUnknown
        # ExSummary:Shows how to work with 'FieldNone' field in a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a field that does not denote an objective field type in its field code.
        field = builder.insert_field(" NOTAREALFIELD #a")

        # The "FieldNone" field type is reserved for fields such as these.
        self.assertEqual(aw.fields.FieldType.FIELD_NONE, field.type)

        # We can also still work with these fields and assign them as instances of the FieldUnknown class.
        field_unknown = field.as_field_unknown()
        self.assertEqual(" NOTAREALFIELD #a", field_unknown.get_field_code())
        # ExEnd

        doc = DocumentHelper.save_open(doc)

        # TestUtil.verify_field(aw.fields.FieldType.FIELD_NONE, " NOTAREALFIELD #a", "Error! Bookmark not defined.", doc.range.fields[0])

    def test_insert_tc_field(self):
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a TC field at the current document builder position.
        builder.insert_field("TC \"Entry Text\" \\f t")

    # def test_insert_tc_fields_at_text(self) :
    #
    #     doc = aw.Document()
    #
    #     options = aw.replacing.FindReplaceOptions()
    #     options.replacing_callback = InsertTcFieldHandler("Chapter 1", "\\l 1")
    #
    #     # Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
    #     doc.range.replace(new Regex("The Beginning"), "", options)
    #
    #
    # class _InsertTcFieldHandler(aw.Irepla):
    #
    #     # Store the text and switches to be used for the TC fields.
    #     _mFieldText = ""
    #     _mFieldSwitches = ""
    #
    #     # <summary>
    #     # The display text and switches to use for each TC field. Display name can be an empty String or null.
    #     # </summary>
    #     public InsertTcFieldHandler(string text, string switches)
    #
    #         mFieldText = text
    #         mFieldSwitches = switches
    #
    #
    #     ReplaceAction IReplacingCallback.replacing(ReplacingArgs args)
    #
    #         DocumentBuilder builder = new DocumentBuilder((Document)args.match_node.document)
    #         builder.move_to(args.match_node)
    #
    #         # If the user-specified text is used in the field as display text, use that, otherwise
    #         # use the match String as the display text.
    #         string insertText = !string.is_null_or_empty(mFieldText) ? mFieldText : args.match.value
    #
    #         # Insert the TC field before this node using the specified String
    #         # as the display text and user-defined switches.
    #         builder.insert_field($"TC \"insertText\" mFieldSwitches")
    #
    #         return ReplaceAction.skip

    @unittest.skip('Working with CultureInfo ??? ')
    def test_field_locale(self):
        # ExStart
        # ExFor:Field.locale_id
        # ExSummary:Shows how to insert a field and work with its locale.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a DATE field, and then print the date it will display.
        # Your thread's current culture determines the formatting of the date.
        field = builder.insert_field("DATE")
        print("Today's date, as displayed in the \"CultureInfo.current_culture.english_name\" culture: field.result")

        self.assertEqual(1033, field.locale_id)
        self.assertEqual(aw.fields.FieldUpdateCultureSource.CURRENT_THREAD,
                         doc.field_options.field_update_culture_source)  # ExSkip

        # Changing the culture of our thread will impact the result of the DATE field.
        # Another way to get the DATE field to display a date in a different culture is to use its LocaleId property.
        # This way allows us to avoid changing the thread's culture to get this effect.
        doc.field_options.field_update_culture_source = aw.fields.FieldUpdateCultureSource.FIELD_CODE
        de = CultureInfo("de-DE")
        field.locale_id = de.lcid
        field.update()

        print(
            "Today's date, as displayed according to the \"CultureInfo.get_culture_info(field.locale_id).english_name\" culture: field.result")
        # ExEnd

        doc = DocumentHelper.save_open(doc)
        field = doc.range.fields[0]

        # TestUtil.verify_field(FieldType.field_date, "DATE", DateTime.now.to_string(de.date_time_format.short_date_pattern), field)
        self.assertEqual(CultureInfo("de-DE").lcid, field.locale_id)

    @unittest.skip("WORDSNET-16037")
    def test_update_dirty_fields(self):

        # ExStart
        # ExFor:Field.is_dirty
        # ExFor:LoadOptions.update_dirty_fields
        # ExSummary:Shows how to use special property for updating field result.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Give the document's built-in "Author" property value, and then display it with a field.
        doc.built_in_document_properties.author = "John Doe"
        field = builder.insert_field(aw.fields.FieldType.field_author, True).as_field_author()

        self.assertFalse(field.is_dirty)
        self.assertEqual("John Doe", field.result)

        # Update the property. The field still displays the old value.
        doc.built_in_document_properties.author = "John & Jane Doe"

        self.assertEqual("John Doe", field.result)

        # Since the field's value is out of date, we can mark it as "dirty".
        # This value will stay out of date until we update the field manually with the Field.update() method.
        field.is_dirty = True

        # docStream = io.BytesIO()
        #
        # # If we save without calling an update method,
        # # the field will keep displaying the out of date value in the output document.
        # doc.save(docStream, aw.saving.SaveFormat.DOCX)
        #
        # # The LoadOptions object has an option to update all fields
        # # marked as "dirty" when loading the document.
        # options = aw.loading.LoadOptions()
        # options.update_dirty_fields = updateDirtyFields
        # doc = new Document(docStream, options)
        #
        # self.assertEqual("John & Jane Doe", doc.built_in_document_properties.author)
        #
        # field = (FieldAuthor)doc.range.fields[0]
        #
        # # Updating dirty fields like this automatically set their "IsDirty" flag to false.
        # if (updateDirtyFields)
        #
        #     self.assertEqual("John & Jane Doe", field.result)
        #     self.assertFalse(field.is_dirty)
        #
        # else
        #
        #     self.assertEqual("John Doe", field.result)
        #     self.assertTrue(field.is_dirty)
        #
        #
        # #ExEnd

    @unittest.skip("raise ArgumentException ???")
    def test_insert_field_with_field_builder_exception(self):

        doc = aw.Document()

        run = DocumentHelper.insert_new_run(DocumentHelper, doc, " Hello World!", 0)

        argument_builder = aw.fields.FieldArgumentBuilder()
        argument_builder.add_field(aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_MERGE_FIELD))
        argument_builder.add_node(run)
        argument_builder.add_text("Text argument builder")

        field_builder = aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_INCLUDE_TEXT)

        # self.assertRaises(RuntimeError, field_builder.add_argument(argument_builder).add_argument("=").add_argument("BestField").add_argument(10).add_argument(20.0).build_and_insert(run))
        # field_builder.add_argument(argument_builder).add_argument("=").add_argument("BestField").add_argument(10).add_argument(20.0).build_and_insert(run)

    # #if NET462 || JAVA
    #     def test_bar_code_word_2_pdf(self) :
    #
    #         Document doc = new Document(aeb.my_dir + "Field sample - BARCODE.docx")
    #
    #         doc.field_options.barcode_generator = new CustomBarcodeGenerator()
    #
    #         doc.save(aeb.artifacts_dir + "Field.bar_code_word_2_pdf.pdf")
    #
    #         using (BarCodeReader barCodeReader = BarCodeReaderPdf(aeb.artifacts_dir + "Field.bar_code_word_2_pdf.pdf"))
    #
    #             self.assertEqual("QR", barCodeReader.found_bar_codes[0].code_type_name)
    #
    #
    #
    #     private BarCodeReader BarCodeReaderPdf(string filename)
    #
    #         # Set license for Aspose.bar_code.
    #         Aspose.bar_code.license licenceBarCode = new Aspose.bar_code.license()
    #         licenceBarCode.set_license(LicenseDir + "Aspose.total.net.lic")
    #
    #         Aspose.pdf.facades.pdf_extractor pdfExtractor = new Aspose.pdf.facades.pdf_extractor()
    #         pdfExtractor.bind_pdf(filename)
    #
    #         # Set page range for image extraction.
    #         pdfExtractor.start_page = 1
    #         pdfExtractor.end_page = 1
    #
    #         pdfExtractor.extract_image()
    #
    #         MemoryStream imageStream = new MemoryStream()
    #         pdfExtractor.get_next_image(imageStream)
    #         imageStream.position = 0
    #
    #         # Recognize the barcode from the image stream above.
    #         BarCodeReader barcodeReader = new BarCodeReader(imageStream, DecodeType.qr)
    #
    #         foreach (BarCodeResult result in barcodeReader.read_bar_codes())
    #             print("Codetext found: " + result.code_text + ", Symbology: " + result.code_type_name)
    #
    #         return barcodeReader
    #
    #
    @unittest.skip("WORDSNET-13854")
    def def_field_database(self):

        # ExStart
        # ExFor:FieldDatabase
        # ExFor:FieldDatabase.connection
        # ExFor:FieldDatabase.file_name
        # ExFor:FieldDatabase.first_record
        # ExFor:FieldDatabase.format_attributes
        # ExFor:FieldDatabase.insert_headings
        # ExFor:FieldDatabase.insert_once_on_mail_merge
        # ExFor:FieldDatabase.last_record
        # ExFor:FieldDatabase.query
        # ExFor:FieldDatabase.table_format
        # ExSummary:Shows how to extract data from a database and insert it as a field into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # This DATABASE field will run a query on a database, and display the result in a table.
        # FieldDatabase field = (FieldDatabase)builder.insert_field(FieldType.field_database, true)
        # field.file_name = aeb.my_dir + @"Database\Northwind.mdb"
        # field.connection = "DSN=MS Access Databases"
        # field.query = "SELECT * FROM [Products]"
        #
        # self.assertEqual($" DATABASE  \\d \"DatabaseDir.replace("\\", "\\\\") + "Northwind.mdb"\" \\c \"DSN=MS Access Databases\" \\s \"SELECT * FROM [Products]\"",
        #     field.get_field_code())
        #
        # # Insert another DATABASE field with a more complex query that sorts all products in descending order by gross sales.
        # field = (FieldDatabase)builder.insert_field(FieldType.field_database, true)
        # field.file_name = aeb.my_dir + @"Database\Northwind.mdb"
        # field.connection = "DSN=MS Access Databases"
        # field.query =
        #     "SELECT [Products].product_name, FORMAT(SUM([Order Details].unit_price * (1 - [Order Details].discount) * [Order Details].quantity), 'Currency') AS GrossSales " +
        #     "FROM([Products] " +
        #     "LEFT JOIN[Order Details] ON[Products].[ProductID] = [Order Details].[ProductID]) " +
        #     "GROUP BY[Products].product_name " +
        #     "ORDER BY SUM([Order Details].unit_price* (1 - [Order Details].discount) * [Order Details].quantity) DESC"
        #
        # # These properties have the same function as LIMIT and TOP clauses.
        # # Configure them to display only rows 1 to 10 of the query result in the field's table.
        # field.first_record = "1"
        # field.last_record = "10"
        #
        # # This property is the index of the format we want to use for our table. The list of table formats is in the "Table AutoFormat..." menu
        # # that shows up when we create a DATABASE field in Microsoft Word. Index #10 corresponds to the "Colorful 3" format.
        # field.table_format = "10"
        #
        # # The FormatAttribute property is a string representation of an integer which stores multiple flags.
        # # We can patrially apply the format which the TableFormat property points to by setting different flags in this property.
        # # The number we use is the sum of a combination of values corresponding to different aspects of the table style.
        # # 63 represents 1 (borders) + 2 (shading) + 4 (font) + 8 (color) + 16 (autofit) + 32 (heading rows).
        # field.format_attributes = "63"
        # field.insert_headings = true
        # field.insert_once_on_mail_merge = true
        #
        # doc.update_fields()
        # doc.save(aeb.artifacts_dir + "Field.database.docx")
        # #ExEnd
        #
        # doc = new Document(aeb.artifacts_dir + "Field.database.docx")
        #
        # self.assertEqual(2, doc.range.fields.count)
        #
        # Table table = doc.first_section.body.tables[0]
        #
        # self.assertEqual(77, table.rows.count)
        # self.assertEqual(10, table.rows[0].cells.count)
        #
        # field = (FieldDatabase)doc.range.fields[0]
        #
        # self.assertEqual($" DATABASE  \\d \"DatabaseDir.replace("\\", "\\\\") + "Northwind.mdb"\" \\c \"DSN=MS Access Databases\" \\s \"SELECT * FROM [Products]\"",
        #     field.get_field_code())
        #
        # TestUtil.table_matches_query_result(table, DatabaseDir + "Northwind.mdb", field.query)
        #
        # table = (Table)doc.get_child(NodeType.table, 1, true)
        # field = (FieldDatabase)doc.range.fields[1]
        #
        # self.assertEqual(11, table.rows.count)
        # self.assertEqual(2, table.rows[0].cells.count)
        # self.assertEqual("ProductName\a", table.rows[0].cells[0].get_text())
        # self.assertEqual("GrossSales\a", table.rows[0].cells[1].get_text())
        #
        # self.assertEqual($" DATABASE  \\d \"DatabaseDir.replace("\\", "\\\\") + "Northwind.mdb"\" \\c \"DSN=MS Access Databases\" " +
        #                 $"\\s \"SELECT [Products].product_name, FORMAT(SUM([Order Details].unit_price * (1 - [Order Details].discount) * [Order Details].quantity), 'Currency') AS GrossSales " +
        #                 "FROM([Products] " +
        #                 "LEFT JOIN[Order Details] ON[Products].[ProductID] = [Order Details].[ProductID]) " +
        #                 "GROUP BY[Products].product_name " +
        #                 "ORDER BY SUM([Order Details].unit_price* (1 - [Order Details].discount) * [Order Details].quantity) DESC\" \\f 1 \\t 10 \\l 10 \\b 63 \\h \\o",
        #     field.get_field_code())
        #
        # table.rows[0].remove()
        #
        # TestUtil.table_matches_query_result(table, DatabaseDir + "Northwind.mdb", field.query.insert(7, " TOP 10 "))

    #
    # #endif

    def test_preserve_include_picture(self):

        # ExStart
        # ExFor:Field.update(bool)
        # ExFor:LoadOptions.preserve_include_picture_field
        # ExSummary:Shows how to preserve or discard INCLUDEPICTURE fields when loading a document.

        for preserve_include_picture_field in (False, True):
            with self.subTest(preserve_include_picture_field=preserve_include_picture_field):
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                include_picture = builder.insert_field(aw.fields.FieldType.FIELD_INCLUDE_PICTURE,
                                                       True).as_field_include_picture()
                include_picture.source_full_name = aeb.image_dir + "Transparent background logo.png"
                include_picture.update(True)

                doc.save(aeb.artifacts_dir + "Fields.preserve_include_picture.docx",
                         aw.saving.OoxmlSaveOptions(aw.SaveFormat.DOCX))

                # We can set a flag in a LoadOptions object to decide whether to convert all INCLUDEPICTURE fields
                # into image shapes when loading a document that contains them.
                load_options = aw.loading.LoadOptions()
                load_options.preserve_include_picture_field = preserve_include_picture_field

                doc = aw.Document(aeb.artifacts_dir + "Fields.preserve_include_picture.docx", load_options)

                if preserve_include_picture_field:

                    self.assertTrue(
                        any(f for f in doc.range.fields if f.type == aw.fields.FieldType.FIELD_INCLUDE_PICTURE))

                    doc.update_fields()
                    doc.save(aeb.artifacts_dir + "Field.preserve_include_picture.docx")

                else:

                    self.assertFalse(
                        any(f for f in doc.range.fields if f.type == aw.fields.FieldType.FIELD_INCLUDE_PICTURE))

        # ExEnd

    def test_field_format(self):

        # ExStart
        # ExFor:Field.format
        # ExFor:Field.update
        # ExFor:FieldFormat
        # ExFor:FieldFormat.date_time_format
        # ExFor:FieldFormat.numeric_format
        # ExFor:FieldFormat.general_formats
        # ExFor:GeneralFormat
        # ExFor:GeneralFormatCollection
        # ExFor:GeneralFormatCollection.add(GeneralFormat)
        # ExFor:GeneralFormatCollection.count
        # ExFor:GeneralFormatCollection.item(Int32)
        # ExFor:GeneralFormatCollection.remove(GeneralFormat)
        # ExFor:GeneralFormatCollection.remove_at(Int32)
        # ExFor:GeneralFormatCollection.get_enumerator
        # ExSummary:Shows how to format field results.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Use a document builder to insert a field that displays a result with no format applied.
        field = builder.insert_field("= 2 + 3")

        self.assertEqual("= 2 + 3", field.get_field_code())
        self.assertEqual("5", field.result)

        # We can apply a format to a field's result using the field's properties.
        # Below are three types of formats that we can apply to a field's result.
        # 1 -  Numeric format:
        format = field.format
        format.numeric_format = "$###.00"
        field.update()

        self.assertEqual("= 2 + 3 \\# $###.00", field.get_field_code())
        self.assertEqual("$  5.00", field.result)

        # 2 -  Date/time format:
        field = builder.insert_field("DATE")
        format = field.format
        format.date_time_format = "dddd, MMMM dd, yyyy"
        field.update()

        self.assertEqual("DATE \\@ \"dddd, MMMM dd, yyyy\"", field.get_field_code())
        print("Today's date, in format.date_time_format format:\n\tfield.result")

        # 3 -  General format:
        field = builder.insert_field("= 25 + 33")
        format = field.format
        format.general_formats.add(aw.fields.GeneralFormat.LOWERCASE_ROMAN)
        format.general_formats.add(aw.fields.GeneralFormat.UPPER)
        field.update()

        index = 0
        for general_format_enumerator in format.general_formats:
            print("General format index " + str(index + 1) + ": " + str(general_format_enumerator))
            index += 1

        self.assertEqual("= 25 + 33 \\* roman \\* Upper", field.get_field_code())
        self.assertEqual("LVIII", field.result)
        self.assertEqual(2, format.general_formats.count)
        self.assertEqual(aw.fields.GeneralFormat.LOWERCASE_ROMAN, format.general_formats[0])

        # We can remove our formats to revert the field's result to its original form.
        format.general_formats.remove(aw.fields.GeneralFormat.LOWERCASE_ROMAN)
        format.general_formats.remove_at(0)
        self.assertEqual(0, format.general_formats.count)
        field.update()

        self.assertEqual("= 25 + 33  ", field.get_field_code())
        self.assertEqual("58", field.result)
        self.assertEqual(0, format.general_formats.count)
        # ExEnd

    def test_unlink(self):

        # ExStart
        # ExFor:Document.unlink_fields
        # ExSummary:Shows how to unlink all fields in the document.
        doc = aw.Document(aeb.my_dir + "Linked fields.docx")

        doc.unlink_fields()
        # ExEnd

        doc = DocumentHelper.save_open(doc)
        para_with_fields = DocumentHelper.get_paragraph_text(doc, 0)

        self.assertEqual("Fields.Docx   Элементы указателя не найдены.     1.\r", para_with_fields)

    def test_unlink_all_fields_in_range(self):

        # ExStart
        # ExFor:Range.unlink_fields
        # ExSummary:Shows how to unlink all fields in a range.
        doc = aw.Document(aeb.my_dir + "Linked fields.docx")

        new_section = doc.sections[0].clone(True).as_section()
        doc.sections.add(new_section)

        doc.sections[1].range.unlink_fields()
        # ExEnd

        doc = DocumentHelper.save_open(doc)
        sec_with_fields = DocumentHelper.get_section_text(doc, 1)

        print(sec_with_fields.strip())
        self.assertTrue(sec_with_fields.strip().endswith(
            "Fields.docx   Элементы указателя не найдены.     3.\rОшибка! Не указана последовательность.    Fields.Docx   Элементы указателя не найдены.     4."))

    def test_unlink_single_field(self):

        # ExStart
        # ExFor:Field.unlink
        # ExSummary:Shows how to unlink a field.
        doc = aw.Document(aeb.my_dir + "Linked fields.docx")
        doc.range.fields[1].unlink()
        # ExEnd

        doc = DocumentHelper.save_open(doc)
        para_with_fields = DocumentHelper.get_paragraph_text(doc, 0)

        print(para_with_fields.strip())
        self.assertTrue(para_with_fields.strip().endswith(
            "FILENAME  \\* Caps  \\* MERGEFORMAT \u0014Fields.docx\u0015   Элементы указателя не найдены.     \u0013 LISTNUM  LegalDefault \u0015"))

    #     def test_update_toc_page_numbers(self) :
    #
    #         Document doc = new Document(aeb.my_dir + "Field sample - TOC.docx")
    #
    #         Node startNode = DocumentHelper.get_paragraph(doc, 2)
    #         Node endNode = null
    #
    #         NodeCollection paragraphCollection = doc.get_child_nodes(NodeType.paragraph, true)
    #
    #         foreach (Paragraph para in paragraphCollection.of_type<Paragraph>())
    #
    #             foreach (Run run in para.runs.of_type<Run>())
    #
    #                 if (run.text.contains(ControlChar.page_break))
    #
    #                     endNode = run
    #                     break
    #
    #
    #
    #
    #         if (startNode != null && endNode != null)
    #
    #             RemoveSequence(startNode, endNode)
    #
    #             startNode.remove()
    #             endNode.remove()
    #
    #
    #         NodeCollection fStart = doc.get_child_nodes(NodeType.field_start, true)
    #
    #         foreach (FieldStart field in fStart.of_type<FieldStart>())
    #
    #             FieldType fType = field.field_type
    #             if (fType == FieldType.field_toc)
    #
    #                 Paragraph para = (Paragraph)field.get_ancestor(NodeType.paragraph)
    #                 para.range.update_fields()
    #                 break
    #
    #
    #
    #         doc.save(aeb.artifacts_dir + "Field.update_toc_page_numbers.docx")
    #
    #
    #     private static void RemoveSequence(Node start, Node end)
    #
    #         Node curNode = start.next_pre_order(start.document)
    #         while (curNode != null && !curNode.equals(end))
    #
    #             Node nextNode = curNode.next_pre_order(start.document)
    #
    #             if (curNode.is_composite)
    #
    #                 CompositeNode curComposite = (CompositeNode)curNode
    #                 if (!curComposite.get_child_nodes(NodeType.any, true).contains(end) &&
    #                     !curComposite.get_child_nodes(NodeType.any, true).contains(start))
    #
    #                     nextNode = curNode.next_sibling
    #                     curNode.remove()
    #
    #
    #             else
    #
    #                 curNode.remove()
    #
    #
    #             curNode = nextNode
    #
    #
    #
    #     #ExStart
    #     #ExFor:Fields.field_ask
    #     #ExFor:Fields.field_ask.bookmark_name
    #     #ExFor:Fields.field_ask.default_response
    #     #ExFor:Fields.field_ask.prompt_once_on_mail_merge
    #     #ExFor:Fields.field_ask.prompt_text
    #     #ExFor:FieldOptions.user_prompt_respondent
    #     #ExFor:IFieldUserPromptRespondent
    #     #ExFor:IFieldUserPromptRespondent.respond(String,String)
    #     #ExSummary:Shows how to create an ASK field, and set its properties.
    #     def test_field_ask(self) :
    #
    #         doc = aw.Document()
    #         builder = aw.DocumentBuilder(doc)
    #
    #         # Place a field where the response to our ASK field will be placed.
    #         FieldRef fieldRef = (FieldRef)builder.insert_field(FieldType.field_ref, true)
    #         fieldRef.bookmark_name = "MyAskField"
    #         builder.writeln()
    #
    #         self.assertEqual(" REF  MyAskField", fieldRef.get_field_code())
    #
    #         # Insert the ASK field and edit its properties to reference our REF field by bookmark name.
    #         FieldAsk fieldAsk = (FieldAsk)builder.insert_field(FieldType.field_ask, true)
    #         fieldAsk.bookmark_name = "MyAskField"
    #         fieldAsk.prompt_text = "Please provide a response for this ASK field"
    #         fieldAsk.default_response = "Response from within the field."
    #         fieldAsk.prompt_once_on_mail_merge = true
    #         builder.writeln()
    #
    #         self.assertEqual(
    #             " ASK  MyAskField \"Please provide a response for this ASK field\" \\d \"Response from within the field.\" \\o",
    #             fieldAsk.get_field_code())
    #
    #         # ASK fields apply the default response to their respective REF fields during a mail merge.
    #         DataTable table = new DataTable("My Table")
    #         table.columns.add("Column 1")
    #         table.rows.add("Row 1")
    #         table.rows.add("Row 2")
    #
    #         FieldMergeField fieldMergeField = (FieldMergeField)builder.insert_field(FieldType.field_merge_field, true)
    #         fieldMergeField.field_name = "Column 1"
    #
    #         # We can modify or override the default response in our ASK fields with a custom prompt responder,
    #         # which will occur during a mail merge.
    #         doc.field_options.user_prompt_respondent = new MyPromptRespondent()
    #         doc.mail_merge.execute(table)
    #
    #         doc.update_fields()
    #         doc.save(aeb.artifacts_dir + "Field.ask.docx")
    #         TestFieldAsk(table, doc) #ExSkip
    #
    #
    #     # <summary>
    #     # Prepends text to the default response of an ASK field during a mail merge.
    #     # </summary>
    #     private class MyPromptRespondent : IFieldUserPromptRespondent
    #
    #         public string Respond(string promptText, string defaultResponse)
    #
    #             return "Response from MyPromptRespondent. " + defaultResponse
    #
    #
    #     #ExEnd
    #
    #     private void TestFieldAsk(DataTable dataTable, Document doc)
    #
    #         doc = DocumentHelper.save_open(doc)
    #
    #         FieldRef fieldRef = (FieldRef)doc.range.fields.first(f => f.type == FieldType.field_ref)
    #         TestUtil.verify_field(FieldType.field_ref,
    #             " REF  MyAskField", "Response from MyPromptRespondent. Response from within the field.", fieldRef)
    #
    #         FieldAsk fieldAsk = (FieldAsk)doc.range.fields.first(f => f.type == FieldType.field_ask)
    #         TestUtil.verify_field(FieldType.field_ask,
    #             " ASK  MyAskField \"Please provide a response for this ASK field\" \\d \"Response from within the field.\" \\o",
    #             "Response from MyPromptRespondent. Response from within the field.", fieldAsk)
    #
    #         self.assertEqual("MyAskField", fieldAsk.bookmark_name)
    #         self.assertEqual("Please provide a response for this ASK field", fieldAsk.prompt_text)
    #         self.assertEqual("Response from within the field.", fieldAsk.default_response)
    #         self.assertEqual(true, fieldAsk.prompt_once_on_mail_merge)
    #
    #         TestUtil.mail_merge_matches_data_table(dataTable, doc, true)
    #
    #
    def test_field_advance(self):

        # ExStart
        # ExFor:Fields.field_advance
        # ExFor:Fields.field_advance.down_offset
        # ExFor:Fields.field_advance.horizontal_position
        # ExFor:Fields.field_advance.left_offset
        # ExFor:Fields.field_advance.right_offset
        # ExFor:Fields.field_advance.up_offset
        # ExFor:Fields.field_advance.vertical_position
        # ExSummary:Shows how to insert an ADVANCE field, and edit its properties.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("This text is in its normal place.")

        # Below are two ways of using the ADVANCE field to adjust the position of text that follows it.
        # The effects of an ADVANCE field continue to be applied until the paragraph ends,
        # or another ADVANCE field updates the offset/coordinate values.
        # 1 -  Specify a directional offset:
        field = builder.insert_field(aw.fields.FieldType.FIELD_ADVANCE, True).as_field_advance()
        self.assertEqual(aw.fields.FieldType.FIELD_ADVANCE, field.type)  # ExSkip
        self.assertEqual(" ADVANCE ", field.get_field_code())  # ExSkip
        field.right_offset = "5"
        field.up_offset = "5"

        self.assertEqual(" ADVANCE  \\r 5 \\u 5", field.get_field_code())

        builder.write("This text will be moved up and to the right.")

        field = builder.insert_field(aw.fields.FieldType.FIELD_ADVANCE, True).as_field_advance()
        field.down_offset = "5"
        field.left_offset = "100"

        self.assertEqual(" ADVANCE  \\d 5 \\l 100", field.get_field_code())

        builder.writeln("This text is moved down and to the left, overlapping the previous text.")

        # 2 -  Move text to a position specified by coordinates:
        field = builder.insert_field(aw.fields.FieldType.FIELD_ADVANCE, True).as_field_advance()
        field.horizontal_position = "-100"
        field.vertical_position = "200"

        self.assertEqual(" ADVANCE  \\x -100 \\y 200", field.get_field_code())

        builder.write("This text is in a custom position.")

        doc.save(aeb.artifacts_dir + "Field.advance.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "Field.advance.docx")

        field = doc.range.fields[0].as_field_advance()

        # TestUtil.verify_field(FieldType.field_advance, " ADVANCE  \\r 5 \\u 5", string.empty, field)
        self.assertEqual("5", field.right_offset)
        self.assertEqual("5", field.up_offset)

        field = doc.range.fields[1].as_field_advance()

        # TestUtil.verify_field(FieldType.field_advance, " ADVANCE  \\d 5 \\l 100", string.empty, field)
        self.assertEqual("5", field.down_offset)
        self.assertEqual("100", field.left_offset)

        field = doc.range.fields[2].as_field_advance()

        # TestUtil.verify_field(FieldType.field_advance, " ADVANCE  \\x -100 \\y 200", string.empty, field)
        self.assertEqual("-100", field.horizontal_position)
        self.assertEqual("200", field.vertical_position)

    @unittest.skip("Working with CultureInfo ??? ")
    def test_field_address_block(self):

        # ExStart
        # ExFor:Fields.field_address_block.excluded_country_or_region_name
        # ExFor:Fields.field_address_block.format_address_on_country_or_region
        # ExFor:Fields.field_address_block.include_country_or_region_name
        # ExFor:Fields.field_address_block.language_id
        # ExFor:Fields.field_address_block.name_and_address_format
        # ExSummary:Shows how to insert an ADDRESSBLOCK field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        field = builder.insert_field(aw.fields.FieldType.FIELD_ADDRESS_BLOCK, True).as_field_address_block()

        self.assertEqual(" ADDRESSBLOCK ", field.get_field_code())

        # Setting this to "2" will include all countries and regions,
        # unless it is the one specified in the ExcludedCountryOrRegionName property.
        field.include_country_or_region_name = "2"
        field.format_address_on_country_or_region = True
        field.excluded_country_or_region_name = "United States"
        field.name_and_address_format = "<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>"

        # By default, this property will contain the language ID of the first character of the document.
        # We can set a different culture for the field to format the result with like this.
        field.language_id = CultureInfo("en-US").lcid.to_string()

        self.assertEqual(
            " ADDRESSBLOCK  \\c 2 \\d \\e \"United States\" \\f \"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>\" \\l 1033",
            field.get_field_code())
        # ExEnd

        doc = DocumentHelper.save_open(doc)
        field = doc.range.fields[0].as_field_address_block()

        # TestUtil.verify_field(FieldType.field_address_block,
        #     " ADDRESSBLOCK  \\c 2 \\d \\e \"United States\" \\f \"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>\" \\l 1033",
        #     "«AddressBlock»", field)
        self.assertEqual("2", field.include_country_or_region_name)
        self.assertEqual(True, field.format_address_on_country_or_region)
        self.assertEqual("United States", field.excluded_country_or_region_name)
        self.assertEqual("<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>",
                         field.name_and_address_format)
        self.assertEqual("1033", field.language_id)

    #     #ExStart
    #     #ExFor:FieldCollection
    #     #ExFor:FieldCollection.count
    #     #ExFor:FieldCollection.get_enumerator
    #     #ExFor:FieldStart
    #     #ExFor:FieldStart.accept(DocumentVisitor)
    #     #ExFor:FieldSeparator
    #     #ExFor:FieldSeparator.accept(DocumentVisitor)
    #     #ExFor:FieldEnd
    #     #ExFor:FieldEnd.accept(DocumentVisitor)
    #     #ExFor:FieldEnd.has_separator
    #     #ExFor:Field.end
    #     #ExFor:Field.separator
    #     #ExFor:Field.start
    #     #ExSummary:Shows how to work with a collection of fields.
    #     def test_field_collection(self) :
    #
    #         doc = aw.Document()
    #         builder = aw.DocumentBuilder(doc)
    #
    #         builder.insert_field(" DATE \\@ \"dddd, d MMMM yyyy\" ")
    #         builder.insert_field(" TIME ")
    #         builder.insert_field(" REVNUM ")
    #         builder.insert_field(" AUTHOR  \"John Doe\" ")
    #         builder.insert_field(" SUBJECT \"My Subject\" ")
    #         builder.insert_field(" QUOTE \"Hello world!\" ")
    #         doc.update_fields()
    #
    #         FieldCollection fields = doc.range.fields
    #
    #         self.assertEqual(6, fields.count)
    #
    #         # Iterate over the field collection, and print contents and type
    #         # of every field using a custom visitor implementation.
    #         FieldVisitor fieldVisitor = new FieldVisitor()
    #
    #         using (IEnumerator<Field> fieldEnumerator = fields.get_enumerator())
    #
    #             while (fieldEnumerator.move_next())
    #
    #                 if (fieldEnumerator.current != null)
    #
    #                     fieldEnumerator.current.start.accept(fieldVisitor)
    #                     fieldEnumerator.current.separator?.accept(fieldVisitor)
    #                     fieldEnumerator.current.end.accept(fieldVisitor)
    #
    #                 else
    #
    #                     print("There are no fields in the document.")
    #
    #
    #
    #
    #         print(fieldVisitor.get_text())
    #         TestFieldCollection(fieldVisitor.get_text()) #ExSkip
    #
    #
    #     # <summary>
    #     # Document visitor implementation that prints field info.
    #     # </summary>
    #     public class FieldVisitor : DocumentVisitor
    #
    #         public FieldVisitor()
    #
    #             mBuilder = new StringBuilder()
    #
    #
    #         # <summary>
    #         # Gets the plain text of the document that was accumulated by the visitor.
    #         # </summary>
    #         public string GetText()
    #
    #             return mBuilder.to_string()
    #
    #
    #         # <summary>
    #         # Called when a FieldStart node is encountered in the document.
    #         # </summary>
    #         public override VisitorAction VisitFieldStart(FieldStart fieldStart)
    #
    #             mBuilder.append_line("Found field: " + fieldStart.field_type)
    #             mBuilder.append_line("\tField code: " + fieldStart.get_field().get_field_code())
    #             mBuilder.append_line("\tDisplayed as: " + fieldStart.get_field().result)
    #
    #             return VisitorAction.continue
    #
    #
    #         # <summary>
    #         # Called when a FieldSeparator node is encountered in the document.
    #         # </summary>
    #         public override VisitorAction VisitFieldSeparator(FieldSeparator fieldSeparator)
    #
    #             mBuilder.append_line("\tFound separator: " + fieldSeparator.get_text())
    #
    #             return VisitorAction.continue
    #
    #
    #         # <summary>
    #         # Called when a FieldEnd node is encountered in the document.
    #         # </summary>
    #         public override VisitorAction VisitFieldEnd(FieldEnd fieldEnd)
    #
    #             mBuilder.append_line("End of field: " + fieldEnd.field_type)
    #
    #             return VisitorAction.continue
    #
    #
    #         private readonly StringBuilder mBuilder
    #
    #     #ExEnd
    #
    #     private void TestFieldCollection(string fieldVisitorText)
    #
    #         self.assertTrue(fieldVisitorText.contains("Found field: FieldDate"))
    #         self.assertTrue(fieldVisitorText.contains("Found field: FieldTime"))
    #         self.assertTrue(fieldVisitorText.contains("Found field: FieldRevisionNum"))
    #         self.assertTrue(fieldVisitorText.contains("Found field: FieldAuthor"))
    #         self.assertTrue(fieldVisitorText.contains("Found field: FieldSubject"))
    #         self.assertTrue(fieldVisitorText.contains("Found field: FieldQuote"))
    #
    #
    def test_remove_fields(self):

        # ExStart
        # ExFor:FieldCollection
        # ExFor:FieldCollection.count
        # ExFor:FieldCollection.clear
        # ExFor:FieldCollection.item(Int32)
        # ExFor:FieldCollection.remove(Field)
        # ExFor:FieldCollection.remove_at(Int32)
        # ExFor:Field.remove
        # ExSummary:Shows how to remove fields from a field collection.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_field(" DATE \\@ \"dddd, d MMMM yyyy\" ")
        builder.insert_field(" TIME ")
        builder.insert_field(" REVNUM ")
        builder.insert_field(" AUTHOR  \"John Doe\" ")
        builder.insert_field(" SUBJECT \"My Subject\" ")
        builder.insert_field(" QUOTE \"Hello world!\" ")
        doc.update_fields()

        fields = doc.range.fields

        self.assertEqual(6, fields.count)

        # Below are four ways of removing fields from a field collection.
        # 1 -  Get a field to remove itself:
        fields[0].remove()
        self.assertEqual(5, fields.count)

        # 2 -  Get the collection to remove a field that we pass to its removal method:
        last_field = fields[3]
        fields.remove(last_field)
        self.assertEqual(4, fields.count)

        # 3 -  Remove a field from a collection at an index:
        fields.remove_at(2)
        self.assertEqual(3, fields.count)

        # 4 -  Remove all the fields from the collection at once:
        fields.clear()
        self.assertEqual(0, fields.count)
        # ExEnd

    def test_field_compare(self):

        # ExStart
        # ExFor:FieldCompare
        # ExFor:FieldCompare.comparison_operator
        # ExFor:FieldCompare.left_expression
        # ExFor:FieldCompare.right_expression
        # ExSummary:Shows how to compare expressions using a COMPARE field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        field = builder.insert_field(aw.fields.FieldType.FIELD_COMPARE, True).as_field_compare()
        field.left_expression = "3"
        field.comparison_operator = "<"
        field.right_expression = "2"
        field.update()

        # The COMPARE field displays a "0" or a "1", depending on its statement's truth.
        # The result of this statement is false so that this field will display a "0".
        self.assertEqual(" COMPARE  3 < 2", field.get_field_code())
        self.assertEqual("0", field.result)

        builder.writeln()

        field = builder.insert_field(aw.fields.FieldType.FIELD_COMPARE, True).as_field_compare()
        field.left_expression = "5"
        field.comparison_operator = "="
        field.right_expression = "2 + 3"
        field.update()

        # This field displays a "1" since the statement is true.
        self.assertEqual(" COMPARE  5 = \"2 + 3\"", field.get_field_code())
        self.assertEqual("1", field.result)

        doc.update_fields()
        doc.save(aeb.artifacts_dir + "Field.compare.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "Field.compare.docx")

        field = doc.range.fields[0].as_field_compare()

        # TestUtil.verify_field(FieldType.field_compare, " COMPARE  3 < 2", "0", field)
        self.assertEqual("3", field.left_expression)
        self.assertEqual("<", field.comparison_operator)
        self.assertEqual("2", field.right_expression)

        field = doc.range.fields[1].as_field_compare()

        # TestUtil.verify_field(FieldType.field_compare, " COMPARE  5 = \"2 + 3\"", "1", field)
        self.assertEqual("5", field.left_expression)
        self.assertEqual("=", field.comparison_operator)
        self.assertEqual("\"2 + 3\"", field.right_expression)

    def test_field_if(self):

        # ExStart
        # ExFor:FieldIf
        # ExFor:FieldIf.comparison_operator
        # ExFor:FieldIf.evaluate_condition
        # ExFor:FieldIf.false_text
        # ExFor:FieldIf.left_expression
        # ExFor:FieldIf.right_expression
        # ExFor:FieldIf.true_text
        # ExFor:FieldIfComparisonResult
        # ExSummary:Shows how to insert an IF field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Statement 1: ")
        field = builder.insert_field(aw.fields.FieldType.FIELD_IF, True).as_field_if()
        field.left_expression = "0"
        field.comparison_operator = "="
        field.right_expression = "1"

        # The IF field will display a string from either its "TrueText" property,
        # or its "FalseText" property, depending on the truth of the statement that we have constructed.
        field.true_text = "True"
        field.false_text = "False"
        field.update()

        # In this case, "0 = 1" is incorrect, so the displayed result will be "False".
        self.assertEqual(" IF  0 = 1 True False", field.get_field_code())
        self.assertEqual(aw.fields.FieldIfComparisonResult.FALSE, field.evaluate_condition())
        self.assertEqual("False", field.result)

        builder.write("\nStatement 2: ")
        field = builder.insert_field(aw.fields.FieldType.FIELD_IF, True).as_field_if()
        field.left_expression = "5"
        field.comparison_operator = "="
        field.right_expression = "2 + 3"
        field.true_text = "True"
        field.false_text = "False"
        field.update()

        # This time the statement is correct, so the displayed result will be "True".
        self.assertEqual(" IF  5 = \"2 + 3\" True False", field.get_field_code())
        self.assertEqual(aw.fields.FieldIfComparisonResult.TRUE, field.evaluate_condition())
        self.assertEqual("True", field.result)

        doc.update_fields()
        doc.save(aeb.artifacts_dir + "Field.if.docx")
        # ExEnd

        doc = aw.Document(aeb.artifacts_dir + "Field.if.docx")
        field = doc.range.fields[0].as_field_if()

        # TestUtil.verify_field(FieldType.field_if, " IF  0 = 1 True False", "False", field)
        self.assertEqual("0", field.left_expression)
        self.assertEqual("=", field.comparison_operator)
        self.assertEqual("1", field.right_expression)
        self.assertEqual("True", field.true_text)
        self.assertEqual("False", field.false_text)

        field = doc.range.fields[1].as_field_if()

        # TestUtil.verify_field(FieldType.field_if, " IF  5 = \"2 + 3\" True False", "True", field)
        self.assertEqual("5", field.left_expression)
        self.assertEqual("=", field.comparison_operator)
        self.assertEqual("\"2 + 3\"", field.right_expression)
        self.assertEqual("True", field.true_text)
        self.assertEqual("False", field.false_text)


    def test_field_auto_num(self) :

        #ExStart
        #ExFor:FieldAutoNum
        #ExFor:FieldAutoNum.separator_character
        #ExSummary:Shows how to number paragraphs using autonum fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Each AUTONUM field displays the current value of a running count of AUTONUM fields,
        # allowing us to automatically number items like a numbered list.
        # This field will display a number "1.".
        field = builder.insert_field(aw.fields.FieldType.FIELD_AUTO_NUM, True).as_field_auto_num()
        builder.writeln("\tParagraph 1.")

        self.assertEqual(" AUTONUM ", field.get_field_code())

        field = builder.insert_field(aw.fields.FieldType.FIELD_AUTO_NUM, True).as_field_auto_num()
        builder.writeln("\tParagraph 2.")

        # The separator character, which appears in the field result immediately after the number,is a full stop by default.
        # If we leave this property null, our second AUTONUM field will display "2." in the document.
        self.assertIsNone(field.separator_character)

        # We can set this property to apply the first character of its string as the new separator character.
        # In this case, our AUTONUM field will now display "2:".
        field.separator_character = ":"

        self.assertEqual(" AUTONUM  \\s :", field.get_field_code())

        doc.save(aeb.artifacts_dir + "Field.autonum.docx")
        #ExEnd

        doc = aw.Document(aeb.artifacts_dir + "Field.autonum.docx")

        # TestUtil.verify_field(FieldType.field_auto_num, " AUTONUM ", string.empty, doc.range.fields[0])
        # TestUtil.verify_field(FieldType.field_auto_num, " AUTONUM  \\s :", string.empty, doc.range.fields[1])


#     #ExStart
#     #ExFor:FieldAutoNumLgl
#     #ExFor:FieldAutoNumLgl.remove_trailing_period
#     #ExFor:FieldAutoNumLgl.separator_character
#     #ExSummary:Shows how to organize a document using AUTONUMLGL fields.
#     def test_field_auto_num_lgl(self) :
#
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         const string fillerText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
#                                   "\nUt enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. "
#
#         # AUTONUMLGL fields display a number that increments at each AUTONUMLGL field within its current heading level.
#         # These fields maintain a separate count for each heading level,
#         # and each field also displays the AUTONUMLGL field counts for all heading levels below its own.
#         # Changing the count for any heading level resets the counts for all levels above that level to 1.
#         # This allows us to organize our document in the form of an outline list.
#         # This is the first AUTONUMLGL field at a heading level of 1, displaying "1." in the document.
#         InsertNumberedClause(builder, "\tHeading 1", fillerText, StyleIdentifier.heading_1)
#
#         # This is the second AUTONUMLGL field at a heading level of 1, so it will display "2.".
#         InsertNumberedClause(builder, "\tHeading 2", fillerText, StyleIdentifier.heading_1)
#
#         # This is the first AUTONUMLGL field at a heading level of 2,
#         # and the AUTONUMLGL count for the heading level below it is "2", so it will display "2.1.".
#         InsertNumberedClause(builder, "\tHeading 3", fillerText, StyleIdentifier.heading_2)
#
#         # This is the first AUTONUMLGL field at a heading level of 3.
#         # Working in the same way as the field above, it will display "2.1.1.".
#         InsertNumberedClause(builder, "\tHeading 4", fillerText, StyleIdentifier.heading_3)
#
#         # This field is at a heading level of 2, and its respective AUTONUMLGL count is at 2, so the field will display "2.2.".
#         InsertNumberedClause(builder, "\tHeading 5", fillerText, StyleIdentifier.heading_2)
#
#         # Incrementing the AUTONUMLGL count for a heading level below this one
#         # has reset the count for this level so that this field will display "2.2.1.".
#         InsertNumberedClause(builder, "\tHeading 6", fillerText, StyleIdentifier.heading_3)
#
#         foreach (FieldAutoNumLgl field in doc.range.fields.where(f => f.type == FieldType.field_auto_num_legal))
#
#             # The separator character, which appears in the field result immediately after the number,
#             # is a full stop by default. If we leave this property null,
#             # our last AUTONUMLGL field will display "2.2.1." in the document.
#             Assert.is_null(field.separator_character)
#
#             # Setting a custom separator character and removing the trailing period
#             # will change that field's appearance from "2.2.1." to "2:2:1".
#             # We will apply this to all the fields that we have created.
#             field.separator_character = ":"
#             field.remove_trailing_period = true
#             self.assertEqual(" AUTONUMLGL  \\s : \\e", field.get_field_code())
#
#
#         doc.save(aeb.artifacts_dir + "Field.autonumlgl.docx")
#         TestFieldAutoNumLgl(doc) #ExSkip
#
#
#     # <summary>
#     # Uses a document builder to insert a clause numbered by an AUTONUMLGL field.
#     # </summary>
#     private static void InsertNumberedClause(DocumentBuilder builder, string heading, string contents, StyleIdentifier headingStyle)
#
#         builder.insert_field(FieldType.field_auto_num_legal, true)
#         builder.current_paragraph.paragraph_format.style_identifier = headingStyle
#         builder.writeln(heading)
#
#         # This text will belong to the auto num legal field above it.
#         # It will collapse when we click the arrow next to the corresponding AUTONUMLGL field in Microsoft Word.
#         builder.current_paragraph.paragraph_format.style_identifier = StyleIdentifier.body_text
#         builder.writeln(contents)
#
#     #ExEnd
#
#     private void TestFieldAutoNumLgl(Document doc)
#
#         doc = DocumentHelper.save_open(doc)
#
#         foreach (FieldAutoNumLgl field in doc.range.fields.where(f => f.type == FieldType.field_auto_num_legal))
#
#             TestUtil.verify_field(FieldType.field_auto_num_legal, " AUTONUMLGL  \\s : \\e", string.empty, field)
#
#             self.assertEqual(":", field.separator_character)
#             self.assertTrue(field.remove_trailing_period)
#
#
#
#     def test_field_auto_num_out(self) :
#
#         #ExStart
#         #ExFor:FieldAutoNumOut
#         #ExSummary:Shows how to number paragraphs using AUTONUMOUT fields.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # AUTONUMOUT fields display a number that increments at each AUTONUMOUT field.
#         # Unlike AUTONUM fields, AUTONUMOUT fields use the outline numbering scheme,
#         # which we can define in Microsoft Word via Format -> Bullets & Numbering -> "Outline Numbered".
#         # This allows us to automatically number items like a numbered list.
#         # LISTNUM fields are a newer alternative to AUTONUMOUT fields.
#         # This field will display "1.".
#         builder.insert_field(FieldType.field_auto_num_outline, true)
#         builder.writeln("\tParagraph 1.")
#
#         # This field will display "2.".
#         builder.insert_field(FieldType.field_auto_num_outline, true)
#         builder.writeln("\tParagraph 2.")
#
#         foreach (FieldAutoNumOut field in doc.range.fields.where(f => f.type == FieldType.field_auto_num_outline))
#             self.assertEqual(" AUTONUMOUT ", field.get_field_code())
#
#         doc.save(aeb.artifacts_dir + "Field.autonumout.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.autonumout.docx")
#
#         foreach (Field field in doc.range.fields)
#             TestUtil.verify_field(FieldType.field_auto_num_outline, " AUTONUMOUT ", string.empty, field)
#
#
#     def test_field_auto_text(self) :
#
#         #ExStart
#         #ExFor:Fields.field_auto_text
#         #ExFor:FieldAutoText.entry_name
#         #ExFor:FieldOptions.built_in_templates_paths
#         #ExFor:FieldGlossary
#         #ExFor:FieldGlossary.entry_name
#         #ExSummary:Shows how to display a building block with AUTOTEXT and GLOSSARY fields.
#         doc = aw.Document()
#
#         # Create a glossary document and add an AutoText building block to it.
#         doc.glossary_document = new GlossaryDocument()
#         BuildingBlock buildingBlock = new BuildingBlock(doc.glossary_document)
#         buildingBlock.name = "MyBlock"
#         buildingBlock.gallery = BuildingBlockGallery.auto_text
#         buildingBlock.category = "General"
#         buildingBlock.description = "MyBlock description"
#         buildingBlock.behavior = BuildingBlockBehavior.paragraph
#         doc.glossary_document.append_child(buildingBlock)
#
#         # Create a source and add it as text to our building block.
#         Document buildingBlockSource = new Document()
#         DocumentBuilder buildingBlockSourceBuilder = new DocumentBuilder(buildingBlockSource)
#         buildingBlockSourceBuilder.writeln("Hello World!")
#
#         Node buildingBlockContent = doc.glossary_document.import_node(buildingBlockSource.first_section, true)
#         buildingBlock.append_child(buildingBlockContent)
#
#         # Set a file which contains parts that our document, or its attached template may not contain.
#         doc.field_options.built_in_templates_paths = new[]  aeb.my_dir + "Busniess brochure.dotx"
#
#         builder = aw.DocumentBuilder(doc)
#
#         # Below are two ways to use fields to display the contents of our building block.
#         # 1 -  Using an AUTOTEXT field:
#         FieldAutoText fieldAutoText = (FieldAutoText)builder.insert_field(FieldType.field_auto_text, true)
#         fieldAutoText.entry_name = "MyBlock"
#
#         self.assertEqual(" AUTOTEXT  MyBlock", fieldAutoText.get_field_code())
#
#         # 2 -  Using a GLOSSARY field:
#         FieldGlossary fieldGlossary = (FieldGlossary)builder.insert_field(FieldType.field_glossary, true)
#         fieldGlossary.entry_name = "MyBlock"
#
#         self.assertEqual(" GLOSSARY  MyBlock", fieldGlossary.get_field_code())
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.autotext.glossary.dotx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.autotext.glossary.dotx")
#
#         Assert.that(doc.field_options.built_in_templates_paths, Is.empty)
#
#         fieldAutoText = (FieldAutoText)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_auto_text, " AUTOTEXT  MyBlock", "Hello World!\r", fieldAutoText)
#         self.assertEqual("MyBlock", fieldAutoText.entry_name)
#
#         fieldGlossary = (FieldGlossary)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_glossary, " GLOSSARY  MyBlock", "Hello World!\r", fieldGlossary)
#         self.assertEqual("MyBlock", fieldGlossary.entry_name)
#
#
#     #ExStart
#     #ExFor:Fields.field_auto_text_list
#     #ExFor:Fields.field_auto_text_list.entry_name
#     #ExFor:Fields.field_auto_text_list.list_style
#     #ExFor:Fields.field_auto_text_list.screen_tip
#     #ExSummary:Shows how to use an AUTOTEXTLIST field to select from a list of AutoText entries.
#     def test_field_auto_text_list(self) :
#
#         doc = aw.Document()
#
#         # Create a glossary document and populate it with auto text entries.
#         doc.glossary_document = new GlossaryDocument()
#         AppendAutoTextEntry(doc.glossary_document, "AutoText 1", "Contents of AutoText 1")
#         AppendAutoTextEntry(doc.glossary_document, "AutoText 2", "Contents of AutoText 2")
#         AppendAutoTextEntry(doc.glossary_document, "AutoText 3", "Contents of AutoText 3")
#
#         builder = aw.DocumentBuilder(doc)
#
#         # Create an AUTOTEXTLIST field and set the text that the field will display in Microsoft Word.
#         # Set the text to prompt the user to right-click this field to select an AutoText building block,
#         # whose contents the field will display.
#         FieldAutoTextList field = (FieldAutoTextList)builder.insert_field(FieldType.field_auto_text_list, true)
#         field.entry_name = "Right click here to select an AutoText block"
#         field.list_style = "Heading 1"
#         field.screen_tip = "Hover tip text for AutoTextList goes here"
#
#         self.assertEqual(" AUTOTEXTLIST  \"Right click here to select an AutoText block\" " +
#                         "\\s \"Heading 1\" " +
#                         "\\t \"Hover tip text for AutoTextList goes here\"", field.get_field_code())
#
#         doc.save(aeb.artifacts_dir + "Field.autotextlist.dotx")
#         TestFieldAutoTextList(doc) #ExSkip
#
#
#     # <summary>
#     # Create an AutoText-type building block and add it to a glossary document.
#     # </summary>
#     private static void AppendAutoTextEntry(GlossaryDocument glossaryDoc, string name, string contents)
#
#         BuildingBlock buildingBlock = new BuildingBlock(glossaryDoc)
#         buildingBlock.name = name
#         buildingBlock.gallery = BuildingBlockGallery.auto_text
#         buildingBlock.category = "General"
#         buildingBlock.behavior = BuildingBlockBehavior.paragraph
#
#         Section section = new Section(glossaryDoc)
#         section.append_child(new Body(glossaryDoc))
#         section.body.append_paragraph(contents)
#         buildingBlock.append_child(section)
#
#         glossaryDoc.append_child(buildingBlock)
#
#     #ExEnd
#
#     private void TestFieldAutoTextList(Document doc)
#
#         doc = DocumentHelper.save_open(doc)
#
#         self.assertEqual(3, doc.glossary_document.count)
#         self.assertEqual("AutoText 1", doc.glossary_document.building_blocks[0].name)
#         self.assertEqual("Contents of AutoText 1", doc.glossary_document.building_blocks[0].get_text().strip())
#         self.assertEqual("AutoText 2", doc.glossary_document.building_blocks[1].name)
#         self.assertEqual("Contents of AutoText 2", doc.glossary_document.building_blocks[1].get_text().strip())
#         self.assertEqual("AutoText 3", doc.glossary_document.building_blocks[2].name)
#         self.assertEqual("Contents of AutoText 3", doc.glossary_document.building_blocks[2].get_text().strip())
#
#         FieldAutoTextList field = (FieldAutoTextList)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_auto_text_list,
#             " AUTOTEXTLIST  \"Right click here to select an AutoText block\" \\s \"Heading 1\" \\t \"Hover tip text for AutoTextList goes here\"",
#             string.empty, field)
#         self.assertEqual("Right click here to select an AutoText block", field.entry_name)
#         self.assertEqual("Heading 1", field.list_style)
#         self.assertEqual("Hover tip text for AutoTextList goes here", field.screen_tip)
#
#
#     def test_field_greeting_line(self) :
#
#         #ExStart
#         #ExFor:FieldGreetingLine
#         #ExFor:FieldGreetingLine.alternate_text
#         #ExFor:FieldGreetingLine.get_field_names
#         #ExFor:FieldGreetingLine.language_id
#         #ExFor:FieldGreetingLine.name_format
#         #ExSummary:Shows how to insert a GREETINGLINE field.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Create a generic greeting using a GREETINGLINE field, and some text after it.
#         FieldGreetingLine field = (FieldGreetingLine)builder.insert_field(FieldType.field_greeting_line, true)
#         builder.writeln("\n\n\tThis is your custom greeting, created programmatically using Aspose Words!")
#
#         # A GREETINGLINE field accepts values from a data source during a mail merge, like a MERGEFIELD.
#         # It can also format how the source's data is written in its place once the mail merge is complete.
#         # The field names collection corresponds to the columns from the data source
#         # that the field will take values from.
#         self.assertEqual(0, field.get_field_names().length)
#
#         # To populate that array, we need to specify a format for our greeting line.
#         field.name_format = "<< _BEFORE_ Dear >><< _TITLE0_ >><< _LAST0_ >><< _AFTER_ ,>> "
#
#         # Now, our field will accept values from these two columns in the data source.
#         self.assertEqual("Courtesy Title", field.get_field_names()[0])
#         self.assertEqual("Last Name", field.get_field_names()[1])
#         self.assertEqual(2, field.get_field_names().length)
#
#         # This string will cover any cases where the data table data is invalid
#         # by substituting the malformed name with a string.
#         field.alternate_text = "Sir or Madam"
#
#         # Set a locale to format the result.
#         field.language_id = new CultureInfo("en-US").lcid.to_string()
#
#         self.assertEqual(" GREETINGLINE  \\f \"<< _BEFORE_ Dear >><< _TITLE0_ >><< _LAST0_ >><< _AFTER_ ,>> \" \\e \"Sir or Madam\" \\l 1033",
#             field.get_field_code())
#
#         # Create a data table with columns whose names match elements
#         # from the field's field names collection, and then carry out the mail merge.
#         DataTable table = new DataTable("Employees")
#         table.columns.add("Courtesy Title")
#         table.columns.add("First Name")
#         table.columns.add("Last Name")
#         table.rows.add("Mr.", "John", "Doe")
#         table.rows.add("Mrs.", "Jane", "Cardholder")
#
#         # This row has an invalid value in the Courtesy Title column, so our greeting will default to the alternate text.
#         table.rows.add("", "No", "Name")
#
#         doc.mail_merge.execute(table)
#
#         Assert.that(doc.range.fields, Is.empty)
#         self.assertEqual("Dear Mr. Doe,\r\r\tThis is your custom greeting, created programmatically using Aspose Words!\r" +
#                         "\fDear Mrs. Cardholder,\r\r\tThis is your custom greeting, created programmatically using Aspose Words!\r" +
#                         "\fDear Sir or Madam,\r\r\tThis is your custom greeting, created programmatically using Aspose Words!",
#             doc.get_text().strip())
#         #ExEnd
#
#
#     def test_field_list_num(self) :
#
#         #ExStart
#         #ExFor:FieldListNum
#         #ExFor:FieldListNum.has_list_name
#         #ExFor:FieldListNum.list_level
#         #ExFor:FieldListNum.list_name
#         #ExFor:FieldListNum.starting_number
#         #ExSummary:Shows how to number paragraphs with LISTNUM fields.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # LISTNUM fields display a number that increments at each LISTNUM field.
#         # These fields also have a variety of options that allow us to use them to emulate numbered lists.
#         FieldListNum field = (FieldListNum)builder.insert_field(FieldType.field_list_num, true)
#
#         # Lists start counting at 1 by default, but we can set this number to a different value, such as 0.
#         # This field will display "0)".
#         field.starting_number = "0"
#         builder.writeln("Paragraph 1")
#
#         self.assertEqual(" LISTNUM  \\s 0", field.get_field_code())
#
#         # LISTNUM fields maintain separate counts for each list level.
#         # Inserting a LISTNUM field in the same paragraph as another LISTNUM field
#         # increases the list level instead of the count.
#         # The next field will continue the count we started above and display a value of "1" at list level 1.
#         builder.insert_field(FieldType.field_list_num, true)
#
#         # This field will start a count at list level 2. It will display a value of "1".
#         builder.insert_field(FieldType.field_list_num, true)
#
#         # This field will start a count at list level 3. It will display a value of "1".
#         # Different list levels have different formatting,
#         # so these fields combined will display a value of "1)a)i)".
#         builder.insert_field(FieldType.field_list_num, true)
#         builder.writeln("Paragraph 2")
#
#         # The next LISTNUM field that we insert will continue the count at the list level
#         # that the previous LISTNUM field was on.
#         # We can use the "ListLevel" property to jump to a different list level.
#         # If this LISTNUM field stayed on list level 3, it would display "ii)",
#         # but, since we have moved it to list level 2, it carries on the count at that level and displays "b)".
#         field = (FieldListNum)builder.insert_field(FieldType.field_list_num, true)
#         field.list_level = "2"
#         builder.writeln("Paragraph 3")
#
#         self.assertEqual(" LISTNUM  \\l 2", field.get_field_code())
#
#         # We can set the ListName property to get the field to emulate a different AUTONUM field type.
#         # "NumberDefault" emulates AUTONUM, "OutlineDefault" emulates AUTONUMOUT,
#         # and "LegalDefault" emulates AUTONUMLGL fields.
#         # The "OutlineDefault" list name with 1 as the starting number will result in displaying "I.".
#         field = (FieldListNum)builder.insert_field(FieldType.field_list_num, true)
#         field.starting_number = "1"
#         field.list_name = "OutlineDefault"
#         builder.writeln("Paragraph 4")
#
#         self.assertTrue(field.has_list_name)
#         self.assertEqual(" LISTNUM  OutlineDefault \\s 1", field.get_field_code())
#
#         # The ListName does not carry over from the previous field, so we will need to set it for each new field.
#         # This field continues the count with the different list name and displays "II.".
#         field = (FieldListNum)builder.insert_field(FieldType.field_list_num, true)
#         field.list_name = "OutlineDefault"
#         builder.writeln("Paragraph 5")
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.listnum.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.listnum.docx")
#
#         self.assertEqual(7, doc.range.fields.count)
#
#         field = (FieldListNum)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_list_num, " LISTNUM  \\s 0", string.empty, field)
#         self.assertEqual("0", field.starting_number)
#         Assert.null(field.list_level)
#         self.assertFalse(field.has_list_name)
#         Assert.null(field.list_name)
#
#         for (int i = 1 i < 4 i++)
#
#             field = (FieldListNum)doc.range.fields[i]
#
#             TestUtil.verify_field(FieldType.field_list_num, " LISTNUM ", string.empty, field)
#             Assert.null(field.starting_number)
#             Assert.null(field.list_level)
#             self.assertFalse(field.has_list_name)
#             Assert.null(field.list_name)
#
#
#         field = (FieldListNum)doc.range.fields[4]
#
#         TestUtil.verify_field(FieldType.field_list_num, " LISTNUM  \\l 2", string.empty, field)
#         Assert.null(field.starting_number)
#         self.assertEqual("2", field.list_level)
#         self.assertFalse(field.has_list_name)
#         Assert.null(field.list_name)
#
#         field = (FieldListNum)doc.range.fields[5]
#
#         TestUtil.verify_field(FieldType.field_list_num, " LISTNUM  OutlineDefault \\s 1", string.empty, field)
#         self.assertEqual("1", field.starting_number)
#         Assert.null(field.list_level)
#         self.assertTrue(field.has_list_name)
#         self.assertEqual("OutlineDefault", field.list_name)
#
#
#     def test_merge_field(self) :
#
#         #ExStart
#         #ExFor:FieldMergeField
#         #ExFor:FieldMergeField.field_name
#         #ExFor:FieldMergeField.field_name_no_prefix
#         #ExFor:FieldMergeField.is_mapped
#         #ExFor:FieldMergeField.is_vertical_formatting
#         #ExFor:FieldMergeField.text_after
#         #ExFor:FieldMergeField.text_before
#         #ExSummary:Shows how to use MERGEFIELD fields to perform a mail merge.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Create a data table to be used as a mail merge data source.
#         DataTable table = new DataTable("Employees")
#         table.columns.add("Courtesy Title")
#         table.columns.add("First Name")
#         table.columns.add("Last Name")
#         table.rows.add("Mr.", "John", "Doe")
#         table.rows.add("Mrs.", "Jane", "Cardholder")
#
#         # Insert a MERGEFIELD with a FieldName property set to the name of a column in the data source.
#         FieldMergeField fieldMergeField = (FieldMergeField)builder.insert_field(FieldType.field_merge_field, true)
#         fieldMergeField.field_name = "Courtesy Title"
#         fieldMergeField.is_mapped = true
#         fieldMergeField.is_vertical_formatting = false
#
#         # We can apply text before and after the value that this field accepts when the merge takes place.
#         fieldMergeField.text_before = "Dear "
#         fieldMergeField.text_after = " "
#
#         self.assertEqual(" MERGEFIELD  \"Courtesy Title\" \\m \\b \"Dear \" \\f \" \"", fieldMergeField.get_field_code())
#
#         # Insert another MERGEFIELD for a different column in the data source.
#         fieldMergeField = (FieldMergeField)builder.insert_field(FieldType.field_merge_field, true)
#         fieldMergeField.field_name = "Last Name"
#         fieldMergeField.text_after = ":"
#
#         doc.update_fields()
#         doc.mail_merge.execute(table)
#
#         self.assertEqual("Dear Mr. Doe:\u000cDear Mrs. Cardholder:", doc.get_text().strip())
#         #ExEnd
#
#         Assert.that(doc.range.fields, Is.empty)
#
#
#     #ExStart
#     #ExFor:FieldToc
#     #ExFor:FieldToc.bookmark_name
#     #ExFor:FieldToc.custom_styles
#     #ExFor:FieldToc.entry_separator
#     #ExFor:FieldToc.heading_level_range
#     #ExFor:FieldToc.hide_in_web_layout
#     #ExFor:FieldToc.insert_hyperlinks
#     #ExFor:FieldToc.page_number_omitting_level_range
#     #ExFor:FieldToc.preserve_line_breaks
#     #ExFor:FieldToc.preserve_tabs
#     #ExFor:FieldToc.update_page_numbers
#     #ExFor:FieldToc.use_paragraph_outline_level
#     #ExFor:FieldOptions.custom_toc_style_separator
#     #ExSummary:Shows how to insert a TOC, and populate it with entries based on heading styles.
#     def test_field_toc(self) :
#
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         builder.start_bookmark("MyBookmark")
#
#         # Insert a TOC field, which will compile all headings into a table of contents.
#         # For each heading, this field will create a line with the text in that heading style to the left,
#         # and the page the heading appears on to the right.
#         FieldToc field = (FieldToc)builder.insert_field(FieldType.field_toc, true)
#
#         # Use the BookmarkName property to only list headings
#         # that appear within the bounds of a bookmark with the "MyBookmark" name.
#         field.bookmark_name = "MyBookmark"
#
#         # Text with a built-in heading style, such as "Heading 1", applied to it will count as a heading.
#         # We can name additional styles to be picked up as headings by the TOC in this property and their TOC levels.
#         field.custom_styles = "Quote 6 Intense Quote 7"
#
#         # By default, Styles/TOC levels are separated in the CustomStyles property by a comma,
#         # but we can set a custom delimiter in this property.
#         doc.field_options.custom_toc_style_separator = ""
#
#         # Configure the field to exclude any headings that have TOC levels outside of this range.
#         field.heading_level_range = "1-3"
#
#         # The TOC will not display the page numbers of headings whose TOC levels are within this range.
#         field.page_number_omitting_level_range = "2-5"
#
#         # Set a custom string that will separate every heading from its page number.
#         field.entry_separator = "-"
#         field.insert_hyperlinks = true
#         field.hide_in_web_layout = false
#         field.preserve_line_breaks = true
#         field.preserve_tabs = true
#         field.use_paragraph_outline_level = false
#
#         InsertNewPageWithHeading(builder, "First entry", "Heading 1")
#         builder.writeln("Paragraph text.")
#         InsertNewPageWithHeading(builder, "Second entry", "Heading 1")
#         InsertNewPageWithHeading(builder, "Third entry", "Quote")
#         InsertNewPageWithHeading(builder, "Fourth entry", "Intense Quote")
#
#         # These two headings will have the page numbers omitted because they are within the "2-5" range.
#         InsertNewPageWithHeading(builder, "Fifth entry", "Heading 2")
#         InsertNewPageWithHeading(builder, "Sixth entry", "Heading 3")
#
#         # This entry does not appear because "Heading 4" is outside of the "1-3" range that we have set earlier.
#         InsertNewPageWithHeading(builder, "Seventh entry", "Heading 4")
#
#         builder.end_bookmark("MyBookmark")
#         builder.writeln("Paragraph text.")
#
#         # This entry does not appear because it is outside the bookmark specified by the TOC.
#         InsertNewPageWithHeading(builder, "Eighth entry", "Heading 1")
#
#         self.assertEqual(" TOC  \\b MyBookmark \\t \"Quote 6 Intense Quote 7\" \\o 1-3 \\n 2-5 \\p - \\h \\x \\w", field.get_field_code())
#
#         field.update_page_numbers()
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.toc.docx")
#         TestFieldToc(doc) #ExSkip
#
#
#     # <summary>
#     # Start a new page and insert a paragraph of a specified style.
#     # </summary>
#     public void InsertNewPageWithHeading(DocumentBuilder builder, string captionText, string styleName)
#
#         builder.insert_break(BreakType.page_break)
#         string originalStyle = builder.paragraph_format.style_name
#         builder.paragraph_format.style = builder.document.styles[styleName]
#         builder.writeln(captionText)
#         builder.paragraph_format.style = builder.document.styles[originalStyle]
#
#     #ExEnd
#
#     private void TestFieldToc(Document doc)
#
#         doc = DocumentHelper.save_open(doc)
#         FieldToc field = (FieldToc)doc.range.fields[0]
#
#         self.assertEqual("MyBookmark", field.bookmark_name)
#         self.assertEqual("Quote 6 Intense Quote 7", field.custom_styles)
#         self.assertEqual("-", field.entry_separator)
#         self.assertEqual("1-3", field.heading_level_range)
#         self.assertEqual("2-5", field.page_number_omitting_level_range)
#         self.assertFalse(field.hide_in_web_layout)
#         self.assertTrue(field.insert_hyperlinks)
#         self.assertTrue(field.preserve_line_breaks)
#         self.assertTrue(field.preserve_tabs)
#         self.assertTrue(field.update_page_numbers())
#         self.assertFalse(field.use_paragraph_outline_level)
#         self.assertEqual(" TOC  \\b MyBookmark \\t \"Quote 6 Intense Quote 7\" \\o 1-3 \\n 2-5 \\p - \\h \\x \\w", field.get_field_code())
#         self.assertEqual("\u0013 HYPERLINK \\l \"_Toc256000001\" \u0014First entry-\u0013 PAGEREF _Toc256000001 \\h \u00142\u0015\u0015\r" +
#                         "\u0013 HYPERLINK \\l \"_Toc256000002\" \u0014Second entry-\u0013 PAGEREF _Toc256000002 \\h \u00143\u0015\u0015\r" +
#                         "\u0013 HYPERLINK \\l \"_Toc256000003\" \u0014Third entry-\u0013 PAGEREF _Toc256000003 \\h \u00144\u0015\u0015\r" +
#                         "\u0013 HYPERLINK \\l \"_Toc256000004\" \u0014Fourth entry-\u0013 PAGEREF _Toc256000004 \\h \u00145\u0015\u0015\r" +
#                         "\u0013 HYPERLINK \\l \"_Toc256000005\" \u0014Fifth entry\u0015\r" +
#                         "\u0013 HYPERLINK \\l \"_Toc256000006\" \u0014Sixth entry\u0015\r", field.result)
#
#
#     #ExStart
#     #ExFor:FieldToc.entry_identifier
#     #ExFor:FieldToc.entry_level_range
#     #ExFor:FieldTC
#     #ExFor:FieldTC.omit_page_number
#     #ExFor:FieldTC.text
#     #ExFor:FieldTC.type_identifier
#     #ExFor:FieldTC.entry_level
#     #ExSummary:Shows how to insert a TOC field, and filter which TC fields end up as entries.
#     def test_field_toc_entry_identifier(self) :
#
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Insert a TOC field, which will compile all TC fields into a table of contents.
#         FieldToc fieldToc = (FieldToc)builder.insert_field(FieldType.field_toc, true)
#
#         # Configure the field only to pick up TC entries of the "A" type, and an entry-level between 1 and 3.
#         fieldToc.entry_identifier = "A"
#         fieldToc.entry_level_range = "1-3"
#
#         self.assertEqual(" TOC  \\f A \\l 1-3", fieldToc.get_field_code())
#
#         # These two entries will appear in the table.
#         builder.insert_break(BreakType.page_break)
#         InsertTocEntry(builder, "TC field 1", "A", "1")
#         InsertTocEntry(builder, "TC field 2", "A", "2")
#
#         self.assertEqual(" TC  \"TC field 1\" \\n \\f A \\l 1", doc.range.fields[1].get_field_code())
#
#         # This entry will be omitted from the table because it has a different type from "A".
#         InsertTocEntry(builder, "TC field 3", "B", "1")
#
#         # This entry will be omitted from the table because it has an entry-level outside of the 1-3 range.
#         InsertTocEntry(builder, "TC field 4", "A", "5")
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.tc.docx")
#         TestFieldTocEntryIdentifier(doc) #ExSkip
#
#
#     # <summary>
#     # Use a document builder to insert a TC field.
#     # </summary>
#     public void InsertTocEntry(DocumentBuilder builder, string text, string typeIdentifier, string entryLevel)
#
#         FieldTC fieldTc = (FieldTC)builder.insert_field(FieldType.field_toc_entry, true)
#         fieldTc.omit_page_number = true
#         fieldTc.text = text
#         fieldTc.type_identifier = typeIdentifier
#         fieldTc.entry_level = entryLevel
#
#     #ExEnd
#
#     private void TestFieldTocEntryIdentifier(Document doc)
#
#         doc = DocumentHelper.save_open(doc)
#         FieldToc fieldToc = (FieldToc)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_toc, " TOC  \\f A \\l 1-3", "TC field 1\rTC field 2\r", fieldToc)
#         self.assertEqual("A", fieldToc.entry_identifier)
#         self.assertEqual("1-3", fieldToc.entry_level_range)
#
#         FieldTC fieldTc = (FieldTC)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_toc_entry, " TC  \"TC field 1\" \\n \\f A \\l 1", string.empty, fieldTc)
#         self.assertTrue(fieldTc.omit_page_number)
#         self.assertEqual("TC field 1", fieldTc.text)
#         self.assertEqual("A", fieldTc.type_identifier)
#         self.assertEqual("1", fieldTc.entry_level)
#
#         fieldTc = (FieldTC)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_toc_entry, " TC  \"TC field 2\" \\n \\f A \\l 2", string.empty, fieldTc)
#         self.assertTrue(fieldTc.omit_page_number)
#         self.assertEqual("TC field 2", fieldTc.text)
#         self.assertEqual("A", fieldTc.type_identifier)
#         self.assertEqual("2", fieldTc.entry_level)
#
#         fieldTc = (FieldTC)doc.range.fields[3]
#
#         TestUtil.verify_field(FieldType.field_toc_entry, " TC  \"TC field 3\" \\n \\f B \\l 1", string.empty, fieldTc)
#         self.assertTrue(fieldTc.omit_page_number)
#         self.assertEqual("TC field 3", fieldTc.text)
#         self.assertEqual("B", fieldTc.type_identifier)
#         self.assertEqual("1", fieldTc.entry_level)
#
#         fieldTc = (FieldTC)doc.range.fields[4]
#
#         TestUtil.verify_field(FieldType.field_toc_entry, " TC  \"TC field 4\" \\n \\f A \\l 5", string.empty, fieldTc)
#         self.assertTrue(fieldTc.omit_page_number)
#         self.assertEqual("TC field 4", fieldTc.text)
#         self.assertEqual("A", fieldTc.type_identifier)
#         self.assertEqual("5", fieldTc.entry_level)
#
#
#     def test_toc_seq_prefix(self) :
#
#         #ExStart
#         #ExFor:FieldToc
#         #ExFor:FieldToc.table_of_figures_label
#         #ExFor:FieldToc.prefixed_sequence_identifier
#         #ExFor:FieldToc.sequence_separator
#         #ExFor:FieldSeq
#         #ExFor:FieldSeq.sequence_identifier
#         #ExSummary:Shows how to populate a TOC field with entries using SEQ fields.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # A TOC field can create an entry in its table of contents for each SEQ field found in the document.
#         # Each entry contains the paragraph that includes the SEQ field and the page's number that the field appears on.
#         FieldToc fieldToc = (FieldToc)builder.insert_field(FieldType.field_toc, true)
#
#         # SEQ fields display a count that increments at each SEQ field.
#         # These fields also maintain separate counts for each unique named sequence
#         # identified by the SEQ field's "SequenceIdentifier" property.
#         # Use the "TableOfFiguresLabel" property to name a main sequence for the TOC.
#         # Now, this TOC will only create entries out of SEQ fields with their "SequenceIdentifier" set to "MySequence".
#         fieldToc.table_of_figures_label = "MySequence"
#
#         # We can name another SEQ field sequence in the "PrefixedSequenceIdentifier" property.
#         # SEQ fields from this prefix sequence will not create TOC entries.
#         # Every TOC entry created from a main sequence SEQ field will now also display the count that
#         # the prefix sequence is currently on at the primary sequence SEQ field that made the entry.
#         fieldToc.prefixed_sequence_identifier = "PrefixSequence"
#
#         # Each TOC entry will display the prefix sequence count immediately to the left
#         # of the page number that the main sequence SEQ field appears on.
#         # We can specify a custom separator that will appear between these two numbers.
#         fieldToc.sequence_separator = ">"
#
#         self.assertEqual(" TOC  \\c MySequence \\s PrefixSequence \\d >", fieldToc.get_field_code())
#
#         builder.insert_break(BreakType.page_break)
#
#         # There are two ways of using SEQ fields to populate this TOC.
#         # 1 -  Inserting a SEQ field that belongs to the TOC's prefix sequence:
#         # This field will increment the SEQ sequence count for the "PrefixSequence" by 1.
#         # Since this field does not belong to the main sequence identified
#         # by the "TableOfFiguresLabel" property of the TOC, it will not appear as an entry.
#         FieldSeq fieldSeq = (FieldSeq)builder.insert_field(FieldType.field_sequence, true)
#         fieldSeq.sequence_identifier = "PrefixSequence"
#         builder.insert_paragraph()
#
#         self.assertEqual(" SEQ  PrefixSequence", fieldSeq.get_field_code())
#
#         # 2 -  Inserting a SEQ field that belongs to the TOC's main sequence:
#         # This SEQ field will create an entry in the TOC.
#         # The TOC entry will contain the paragraph that the SEQ field is in and the number of the page that it appears on.
#         # This entry will also display the count that the prefix sequence is currently at,
#         # separated from the page number by the value in the TOC's SeqenceSeparator property.
#         # The "PrefixSequence" count is at 1, this main sequence SEQ field is on page 2,
#         # and the separator is ">", so entry will display "1>2".
#         builder.write("First TOC entry, MySequence #")
#         fieldSeq = (FieldSeq)builder.insert_field(FieldType.field_sequence, true)
#         fieldSeq.sequence_identifier = "MySequence"
#
#         self.assertEqual(" SEQ  MySequence", fieldSeq.get_field_code())
#
#         # Insert a page, advance the prefix sequence by 2, and insert a SEQ field to create a TOC entry afterwards.
#         # The prefix sequence is now at 2, and the main sequence SEQ field is on page 3,
#         # so the TOC entry will display "2>3" at its page count.
#         builder.insert_break(BreakType.page_break)
#         fieldSeq = (FieldSeq)builder.insert_field(FieldType.field_sequence, true)
#         fieldSeq.sequence_identifier = "PrefixSequence"
#         builder.insert_paragraph()
#         fieldSeq = (FieldSeq)builder.insert_field(FieldType.field_sequence, true)
#         builder.write("Second TOC entry, MySequence #")
#         fieldSeq.sequence_identifier = "MySequence"
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.toc.seq.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.toc.seq.docx")
#
#         self.assertEqual(9, doc.range.fields.count)
#
#         fieldToc = (FieldToc)doc.range.fields[0]
#         print(fieldToc.display_result)
#         TestUtil.verify_field(FieldType.field_toc, " TOC  \\c MySequence \\s PrefixSequence \\d >",
#             "First TOC entry, MySequence #12\t\u0013 SEQ PrefixSequence _Toc256000000 \\* ARABIC \u00141\u0015>\u0013 PAGEREF _Toc256000000 \\h \u00142\u0015\r2" +
#             "Second TOC entry, MySequence #\t\u0013 SEQ PrefixSequence _Toc256000001 \\* ARABIC \u00142\u0015>\u0013 PAGEREF _Toc256000001 \\h \u00143\u0015\r",
#             fieldToc)
#         self.assertEqual("MySequence", fieldToc.table_of_figures_label)
#         self.assertEqual("PrefixSequence", fieldToc.prefixed_sequence_identifier)
#         self.assertEqual(">", fieldToc.sequence_separator)
#
#         fieldSeq = (FieldSeq)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_sequence, " SEQ PrefixSequence _Toc256000000 \\* ARABIC ", "1", fieldSeq)
#         self.assertEqual("PrefixSequence", fieldSeq.sequence_identifier)
#
#         # Byproduct field created by Aspose.words
#         FieldPageRef fieldPageRef = (FieldPageRef)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_page_ref, " PAGEREF _Toc256000000 \\h ", "2", fieldPageRef)
#         self.assertEqual("PrefixSequence", fieldSeq.sequence_identifier)
#         self.assertEqual("_Toc256000000", fieldPageRef.bookmark_name)
#
#         fieldSeq = (FieldSeq)doc.range.fields[3]
#
#         TestUtil.verify_field(FieldType.field_sequence, " SEQ PrefixSequence _Toc256000001 \\* ARABIC ", "2", fieldSeq)
#         self.assertEqual("PrefixSequence", fieldSeq.sequence_identifier)
#
#         fieldPageRef = (FieldPageRef)doc.range.fields[4]
#
#         TestUtil.verify_field(FieldType.field_page_ref, " PAGEREF _Toc256000001 \\h ", "3", fieldPageRef)
#         self.assertEqual("PrefixSequence", fieldSeq.sequence_identifier)
#         self.assertEqual("_Toc256000001", fieldPageRef.bookmark_name)
#
#         fieldSeq = (FieldSeq)doc.range.fields[5]
#
#         TestUtil.verify_field(FieldType.field_sequence, " SEQ  PrefixSequence", "1", fieldSeq)
#         self.assertEqual("PrefixSequence", fieldSeq.sequence_identifier)
#
#         fieldSeq = (FieldSeq)doc.range.fields[6]
#
#         TestUtil.verify_field(FieldType.field_sequence, " SEQ  MySequence", "1", fieldSeq)
#         self.assertEqual("MySequence", fieldSeq.sequence_identifier)
#
#         fieldSeq = (FieldSeq)doc.range.fields[7]
#
#         TestUtil.verify_field(FieldType.field_sequence, " SEQ  PrefixSequence", "2", fieldSeq)
#         self.assertEqual("PrefixSequence", fieldSeq.sequence_identifier)
#
#         fieldSeq = (FieldSeq)doc.range.fields[8]
#
#         TestUtil.verify_field(FieldType.field_sequence, " SEQ  MySequence", "2", fieldSeq)
#         self.assertEqual("MySequence", fieldSeq.sequence_identifier)
#
#
#     def test_toc_seq_numbering(self) :
#
#         #ExStart
#         #ExFor:FieldSeq
#         #ExFor:FieldSeq.insert_next_number
#         #ExFor:FieldSeq.reset_heading_level
#         #ExFor:FieldSeq.reset_number
#         #ExFor:FieldSeq.sequence_identifier
#         #ExSummary:Shows create numbering using SEQ fields.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # SEQ fields display a count that increments at each SEQ field.
#         # These fields also maintain separate counts for each unique named sequence
#         # identified by the SEQ field's "SequenceIdentifier" property.
#         # Insert a SEQ field that will display the current count value of "MySequence",
#         # after using the "ResetNumber" property to set it to 100.
#         builder.write("#")
#         FieldSeq fieldSeq = (FieldSeq)builder.insert_field(FieldType.field_sequence, true)
#         fieldSeq.sequence_identifier = "MySequence"
#         fieldSeq.reset_number = "100"
#         fieldSeq.update()
#
#         self.assertEqual(" SEQ  MySequence \\r 100", fieldSeq.get_field_code())
#         self.assertEqual("100", fieldSeq.result)
#
#         # Display the next number in this sequence with another SEQ field.
#         builder.write(", #")
#         fieldSeq = (FieldSeq)builder.insert_field(FieldType.field_sequence, true)
#         fieldSeq.sequence_identifier = "MySequence"
#         fieldSeq.update()
#
#         self.assertEqual("101", fieldSeq.result)
#
#         # Insert a level 1 heading.
#         builder.insert_break(BreakType.paragraph_break)
#         builder.paragraph_format.style = doc.styles["Heading 1"]
#         builder.writeln("This level 1 heading will reset MySequence to 1")
#         builder.paragraph_format.style = doc.styles["Normal"]
#
#         # Insert another SEQ field from the same sequence and configure it to reset the count at every heading with 1.
#         builder.write("\n#")
#         fieldSeq = (FieldSeq)builder.insert_field(FieldType.field_sequence, true)
#         fieldSeq.sequence_identifier = "MySequence"
#         fieldSeq.reset_heading_level = "1"
#         fieldSeq.update()
#
#         # The above heading is a level 1 heading, so the count for this sequence is reset to 1.
#         self.assertEqual(" SEQ  MySequence \\s 1", fieldSeq.get_field_code())
#         self.assertEqual("1", fieldSeq.result)
#
#         # Move to the next number of this sequence.
#         builder.write(", #")
#         fieldSeq = (FieldSeq)builder.insert_field(FieldType.field_sequence, true)
#         fieldSeq.sequence_identifier = "MySequence"
#         fieldSeq.insert_next_number = true
#         fieldSeq.update()
#
#         self.assertEqual(" SEQ  MySequence \\n", fieldSeq.get_field_code())
#         self.assertEqual("2", fieldSeq.result)
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.seq.reset_numbering.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.seq.reset_numbering.docx")
#
#         self.assertEqual(4, doc.range.fields.count)
#
#         fieldSeq = (FieldSeq)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_sequence, " SEQ  MySequence \\r 100", "100", fieldSeq)
#         self.assertEqual("MySequence", fieldSeq.sequence_identifier)
#
#         fieldSeq = (FieldSeq)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_sequence, " SEQ  MySequence", "101", fieldSeq)
#         self.assertEqual("MySequence", fieldSeq.sequence_identifier)
#
#         fieldSeq = (FieldSeq)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_sequence, " SEQ  MySequence \\s 1", "1", fieldSeq)
#         self.assertEqual("MySequence", fieldSeq.sequence_identifier)
#
#         fieldSeq = (FieldSeq)doc.range.fields[3]
#
#         TestUtil.verify_field(FieldType.field_sequence, " SEQ  MySequence \\n", "2", fieldSeq)
#         self.assertEqual("MySequence", fieldSeq.sequence_identifier)
#
#
#     [Test]
#     [Ignore("WORDSNET-18083")]
#     public void TocSeqBookmark()
#
#         #ExStart
#         #ExFor:FieldSeq
#         #ExFor:FieldSeq.bookmark_name
#         #ExSummary:Shows how to combine table of contents and sequence fields.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # A TOC field can create an entry in its table of contents for each SEQ field found in the document.
#         # Each entry contains the paragraph that contains the SEQ field,
#         # and the number of the page that the field appears on.
#         FieldToc fieldToc = (FieldToc)builder.insert_field(FieldType.field_toc, true)
#
#         # Configure this TOC field to have a SequenceIdentifier property with a value of "MySequence".
#         fieldToc.table_of_figures_label = "MySequence"
#
#         # Configure this TOC field to only pick up SEQ fields that are within the bounds of a bookmark
#         # named "TOCBookmark".
#         fieldToc.bookmark_name = "TOCBookmark"
#         builder.insert_break(BreakType.page_break)
#
#         self.assertEqual(" TOC  \\c MySequence \\b TOCBookmark", fieldToc.get_field_code())
#
#         # SEQ fields display a count that increments at each SEQ field.
#         # These fields also maintain separate counts for each unique named sequence
#         # identified by the SEQ field's "SequenceIdentifier" property.
#         # Insert a SEQ field that has a sequence identifier that matches the TOC's
#         # TableOfFiguresLabel property. This field will not create an entry in the TOC since it is outside
#         # the bookmark's bounds designated by "BookmarkName".
#         builder.write("MySequence #")
#         FieldSeq fieldSeq = (FieldSeq)builder.insert_field(FieldType.field_sequence, true)
#         fieldSeq.sequence_identifier = "MySequence"
#         builder.writeln(", will not show up in the TOC because it is outside of the bookmark.")
#
#         builder.start_bookmark("TOCBookmark")
#
#         # This SEQ field's sequence matches the TOC's "TableOfFiguresLabel" property and is within the bookmark's bounds.
#         # The paragraph that contains this field will show up in the TOC as an entry.
#         builder.write("MySequence #")
#         fieldSeq = (FieldSeq)builder.insert_field(FieldType.field_sequence, true)
#         fieldSeq.sequence_identifier = "MySequence"
#         builder.writeln(", will show up in the TOC next to the entry for the above caption.")
#
#         # This SEQ field's sequence does not match the TOC's "TableOfFiguresLabel" property,
#         # and is within the bounds of the bookmark. Its paragraph will not show up in the TOC as an entry.
#         builder.write("MySequence #")
#         fieldSeq = (FieldSeq)builder.insert_field(FieldType.field_sequence, true)
#         fieldSeq.sequence_identifier = "OtherSequence"
#         builder.writeln(", will not show up in the TOC because it's from a different sequence identifier.")
#
#         # This SEQ field's sequence matches the TOC's "TableOfFiguresLabel" property and is within the bounds of the bookmark.
#         # This field also references another bookmark. The contents of that bookmark will appear in the TOC entry for this SEQ field.
#         # The SEQ field itself will not display the contents of that bookmark.
#         fieldSeq = (FieldSeq)builder.insert_field(FieldType.field_sequence, true)
#         fieldSeq.sequence_identifier = "MySequence"
#         fieldSeq.bookmark_name = "SEQBookmark"
#         self.assertEqual(" SEQ  MySequence SEQBookmark", fieldSeq.get_field_code())
#
#         # Create a bookmark with contents that will show up in the TOC entry due to the above SEQ field referencing it.
#         builder.insert_break(BreakType.page_break)
#         builder.start_bookmark("SEQBookmark")
#         builder.write("MySequence #")
#         fieldSeq = (FieldSeq)builder.insert_field(FieldType.field_sequence, true)
#         fieldSeq.sequence_identifier = "MySequence"
#         builder.writeln(", text from inside SEQBookmark.")
#         builder.end_bookmark("SEQBookmark")
#
#         builder.end_bookmark("TOCBookmark")
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.seq.bookmark.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.seq.bookmark.docx")
#
#         self.assertEqual(8, doc.range.fields.count)
#
#         fieldToc = (FieldToc)doc.range.fields[0]
#         string[] pageRefIds = fieldToc.result.split(' ').where(s => s.starts_with("_Toc")).to_array()
#
#         self.assertEqual(FieldType.field_toc, fieldToc.type)
#         self.assertEqual("MySequence", fieldToc.table_of_figures_label)
#         TestUtil.verify_field(FieldType.field_toc, " TOC  \\c MySequence \\b TOCBookmark",
#             $"MySequence #2, will show up in the TOC next to the entry for the above caption.\t\u0013 PAGEREF pageRefIds[0] \\h \u00142\u0015\r" +
#             $"3MySequence #3, text from inside SEQBookmark.\t\u0013 PAGEREF pageRefIds[1] \\h \u00142\u0015\r", fieldToc)
#
#         FieldPageRef fieldPageRef = (FieldPageRef)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_page_ref, $" PAGEREF pageRefIds[0] \\h ", "2", fieldPageRef)
#         self.assertEqual(pageRefIds[0], fieldPageRef.bookmark_name)
#
#         fieldPageRef = (FieldPageRef)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_page_ref, $" PAGEREF pageRefIds[1] \\h ", "2", fieldPageRef)
#         self.assertEqual(pageRefIds[1], fieldPageRef.bookmark_name)
#
#         fieldSeq = (FieldSeq)doc.range.fields[3]
#
#         TestUtil.verify_field(FieldType.field_sequence, " SEQ  MySequence", "1", fieldSeq)
#         self.assertEqual("MySequence", fieldSeq.sequence_identifier)
#
#         fieldSeq = (FieldSeq)doc.range.fields[4]
#
#         TestUtil.verify_field(FieldType.field_sequence, " SEQ  MySequence", "2", fieldSeq)
#         self.assertEqual("MySequence", fieldSeq.sequence_identifier)
#
#         fieldSeq = (FieldSeq)doc.range.fields[5]
#
#         TestUtil.verify_field(FieldType.field_sequence, " SEQ  OtherSequence", "1", fieldSeq)
#         self.assertEqual("OtherSequence", fieldSeq.sequence_identifier)
#
#         fieldSeq = (FieldSeq)doc.range.fields[6]
#
#         TestUtil.verify_field(FieldType.field_sequence, " SEQ  MySequence SEQBookmark", "3", fieldSeq)
#         self.assertEqual("MySequence", fieldSeq.sequence_identifier)
#         self.assertEqual("SEQBookmark", fieldSeq.bookmark_name)
#
#         fieldSeq = (FieldSeq)doc.range.fields[7]
#
#         TestUtil.verify_field(FieldType.field_sequence, " SEQ  MySequence", "3", fieldSeq)
#         self.assertEqual("MySequence", fieldSeq.sequence_identifier)
#
#
#     [Test]
#     [Ignore("WORDSNET-13854")]
#     public void FieldCitation()
#
#         #ExStart
#         #ExFor:FieldCitation
#         #ExFor:FieldCitation.another_source_tag
#         #ExFor:FieldCitation.format_language_id
#         #ExFor:FieldCitation.page_number
#         #ExFor:FieldCitation.prefix
#         #ExFor:FieldCitation.source_tag
#         #ExFor:FieldCitation.suffix
#         #ExFor:FieldCitation.suppress_author
#         #ExFor:FieldCitation.suppress_title
#         #ExFor:FieldCitation.suppress_year
#         #ExFor:FieldCitation.volume_number
#         #ExFor:FieldBibliography
#         #ExFor:FieldBibliography.format_language_id
#         #ExSummary:Shows how to work with CITATION and BIBLIOGRAPHY fields.
#         # Open a document containing bibliographical sources that we can find in
#         # Microsoft Word via References -> Citations & Bibliography -> Manage Sources.
#         Document doc = new Document(aeb.my_dir + "Bibliography.docx")
#         self.assertEqual(2, doc.range.fields.count) #ExSkip
#
#         builder = aw.DocumentBuilder(doc)
#         builder.write("Text to be cited with one source.")
#
#         # Create a citation with just the page number and the author of the referenced book.
#         FieldCitation fieldCitation = (FieldCitation)builder.insert_field(FieldType.field_citation, true)
#
#         # We refer to sources using their tag names.
#         fieldCitation.source_tag = "Book1"
#         fieldCitation.page_number = "85"
#         fieldCitation.suppress_author = false
#         fieldCitation.suppress_title = true
#         fieldCitation.suppress_year = true
#
#         self.assertEqual(" CITATION  Book1 \\p 85 \\t \\y", fieldCitation.get_field_code())
#
#         # Create a more detailed citation which cites two sources.
#         builder.insert_paragraph()
#         builder.write("Text to be cited with two sources.")
#         fieldCitation = (FieldCitation)builder.insert_field(FieldType.field_citation, true)
#         fieldCitation.source_tag = "Book1"
#         fieldCitation.another_source_tag = "Book2"
#         fieldCitation.format_language_id = "en-US"
#         fieldCitation.page_number = "19"
#         fieldCitation.prefix = "Prefix "
#         fieldCitation.suffix = " Suffix"
#         fieldCitation.suppress_author = false
#         fieldCitation.suppress_title = false
#         fieldCitation.suppress_year = false
#         fieldCitation.volume_number = "VII"
#
#         self.assertEqual(" CITATION  Book1 \\m Book2 \\l en-US \\p 19 \\f \"Prefix \" \\s \" Suffix\" \\v VII", fieldCitation.get_field_code())
#
#         # We can use a BIBLIOGRAPHY field to display all the sources within the document.
#         builder.insert_break(BreakType.page_break)
#         FieldBibliography fieldBibliography = (FieldBibliography)builder.insert_field(FieldType.field_bibliography, true)
#         fieldBibliography.format_language_id = "1124"
#
#         self.assertEqual(" BIBLIOGRAPHY  \\l 1124", fieldBibliography.get_field_code())
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.citation.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.citation.docx")
#
#         self.assertEqual(5, doc.range.fields.count)
#
#         fieldCitation = (FieldCitation)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_citation, " CITATION  Book1 \\p 85 \\t \\y", " (Doe, p. 85)", fieldCitation)
#         self.assertEqual("Book1", fieldCitation.source_tag)
#         self.assertEqual("85", fieldCitation.page_number)
#         self.assertFalse(fieldCitation.suppress_author)
#         self.assertTrue(fieldCitation.suppress_title)
#         self.assertTrue(fieldCitation.suppress_year)
#
#         fieldCitation = (FieldCitation)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_citation,
#             " CITATION  Book1 \\m Book2 \\l en-US \\p 19 \\f \"Prefix \" \\s \" Suffix\" \\v VII",
#             " (Doe, 2018 Prefix Cardholder, 2018, VII:19 Suffix)", fieldCitation)
#         self.assertEqual("Book1", fieldCitation.source_tag)
#         self.assertEqual("Book2", fieldCitation.another_source_tag)
#         self.assertEqual("en-US", fieldCitation.format_language_id)
#         self.assertEqual("Prefix ", fieldCitation.prefix)
#         self.assertEqual(" Suffix", fieldCitation.suffix)
#         self.assertEqual("19", fieldCitation.page_number)
#         self.assertFalse(fieldCitation.suppress_author)
#         self.assertFalse(fieldCitation.suppress_title)
#         self.assertFalse(fieldCitation.suppress_year)
#         self.assertEqual("VII", fieldCitation.volume_number)
#
#         fieldBibliography = (FieldBibliography)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_bibliography, " BIBLIOGRAPHY  \\l 1124",
#             "Cardholder, A. (2018). My Book, Vol. II. New York: Doe Co. Ltd.\rDoe, J. (2018). My Book, Vol I. London: Doe Co. Ltd.\r", fieldBibliography)
#         self.assertEqual("1124", fieldBibliography.format_language_id)
#
#         fieldCitation = (FieldCitation)doc.range.fields[3]
#
#         TestUtil.verify_field(FieldType.field_citation, " CITATION Book1 \\l 1033 ", "(Doe, 2018)", fieldCitation)
#         self.assertEqual("Book1", fieldCitation.source_tag)
#         self.assertEqual("1033", fieldCitation.format_language_id)
#
#         fieldBibliography = (FieldBibliography)doc.range.fields[4]
#
#         TestUtil.verify_field(FieldType.field_bibliography, " BIBLIOGRAPHY ",
#             "Cardholder, A. (2018). My Book, Vol. II. New York: Doe Co. Ltd.\rDoe, J. (2018). My Book, Vol I. London: Doe Co. Ltd.\r", fieldBibliography)
#
#
#     def test_field_data(self) :
#
#         #ExStart
#         #ExFor:FieldData
#         #ExSummary:Shows how to insert a DATA field into a document.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         FieldData field = (FieldData)builder.insert_field(FieldType.field_data, true)
#         self.assertEqual(" DATA ", field.get_field_code())
#         #ExEnd
#
#         TestUtil.verify_field(FieldType.field_data, " DATA ", string.empty, DocumentHelper.save_open(doc).range.fields[0])
#
#
#     def test_field_include(self) :
#
#         #ExStart
#         #ExFor:FieldInclude
#         #ExFor:FieldInclude.bookmark_name
#         #ExFor:FieldInclude.lock_fields
#         #ExFor:FieldInclude.source_full_name
#         #ExFor:FieldInclude.text_converter
#         #ExSummary:Shows how to create an INCLUDE field, and set its properties.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # We can use an INCLUDE field to import a portion of another document in the local file system.
#         # The bookmark from the other document that we reference with this field contains this imported portion.
#         FieldInclude field = (FieldInclude)builder.insert_field(FieldType.field_include, true)
#         field.source_full_name = aeb.my_dir + "Bookmarks.docx"
#         field.bookmark_name = "MyBookmark1"
#         field.lock_fields = false
#         field.text_converter = "Microsoft Word"
#
#         self.assertTrue(Regex.match(field.get_field_code(), " INCLUDE .* MyBookmark1 \\\\c \"Microsoft Word\"").success)
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.include.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.include.docx")
#         field = (FieldInclude)doc.range.fields[0]
#
#         self.assertEqual(FieldType.field_include, field.type)
#         self.assertEqual("First bookmark.", field.result)
#         self.assertTrue(Regex.match(field.get_field_code(), " INCLUDE .* MyBookmark1 \\\\c \"Microsoft Word\"").success)
#
#         self.assertEqual(aeb.my_dir + "Bookmarks.docx", field.source_full_name)
#         self.assertEqual("MyBookmark1", field.bookmark_name)
#         self.assertFalse(field.lock_fields)
#         self.assertEqual("Microsoft Word", field.text_converter)
#
#
#     def test_field_include_picture(self) :
#
#         #ExStart
#         #ExFor:FieldIncludePicture
#         #ExFor:FieldIncludePicture.graphic_filter
#         #ExFor:FieldIncludePicture.is_linked
#         #ExFor:FieldIncludePicture.resize_horizontally
#         #ExFor:FieldIncludePicture.resize_vertically
#         #ExFor:FieldIncludePicture.source_full_name
#         #ExFor:FieldImport
#         #ExFor:FieldImport.graphic_filter
#         #ExFor:FieldImport.is_linked
#         #ExFor:FieldImport.source_full_name
#         #ExSummary:Shows how to insert images using IMPORT and INCLUDEPICTURE fields.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Below are two similar field types that we can use to display images linked from the local file system.
#         # 1 -  The INCLUDEPICTURE field:
#         FieldIncludePicture fieldIncludePicture = (FieldIncludePicture)builder.insert_field(FieldType.field_include_picture, true)
#         fieldIncludePicture.source_full_name = aeb.image_dir + "Transparent background logo.png"
#
#         self.assertTrue(Regex.match(fieldIncludePicture.get_field_code(), " INCLUDEPICTURE  .*").success)
#
#         # Apply the PNG32.flt filter.
#         fieldIncludePicture.graphic_filter = "PNG32"
#         fieldIncludePicture.is_linked = true
#         fieldIncludePicture.resize_horizontally = true
#         fieldIncludePicture.resize_vertically = true
#
#         # 2 -  The IMPORT field:
#         FieldImport fieldImport = (FieldImport)builder.insert_field(FieldType.field_import, true)
#         fieldImport.source_full_name = aeb.image_dir + "Transparent background logo.png"
#         fieldImport.graphic_filter = "PNG32"
#         fieldImport.is_linked = true
#
#         self.assertTrue(Regex.match(fieldImport.get_field_code(), " IMPORT  .* \\\\c PNG32 \\\\d").success)
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.import.includepicture.docx")
#         #ExEnd
#
#         self.assertEqual(aeb.image_dir + "Transparent background logo.png", fieldIncludePicture.source_full_name)
#         self.assertEqual("PNG32", fieldIncludePicture.graphic_filter)
#         self.assertTrue(fieldIncludePicture.is_linked)
#         self.assertTrue(fieldIncludePicture.resize_horizontally)
#         self.assertTrue(fieldIncludePicture.resize_vertically)
#
#         self.assertEqual(aeb.image_dir + "Transparent background logo.png", fieldImport.source_full_name)
#         self.assertEqual("PNG32", fieldImport.graphic_filter)
#         self.assertTrue(fieldImport.is_linked)
#
#         doc = new Document(aeb.artifacts_dir + "Field.import.includepicture.docx")
#
#         # The INCLUDEPICTURE fields have been converted into shapes with linked images during loading.
#         self.assertEqual(0, doc.range.fields.count)
#         self.assertEqual(2, doc.get_child_nodes(NodeType.shape, true).count)
#
#         Shape image = (Shape)doc.get_child(NodeType.shape, 0, true)
#
#         self.assertTrue(image.is_image)
#         Assert.null(image.image_data.image_bytes)
#         self.assertEqual(aeb.image_dir + "Transparent background logo.png", image.image_data.source_full_name.replace("%20", " "))
#
#         image = (Shape)doc.get_child(NodeType.shape, 1, true)
#
#         self.assertTrue(image.is_image)
#         Assert.null(image.image_data.image_bytes)
#         self.assertEqual(aeb.image_dir + "Transparent background logo.png", image.image_data.source_full_name.replace("%20", " "))
#
#
#     #ExStart
#     #ExFor:FieldIncludeText
#     #ExFor:FieldIncludeText.bookmark_name
#     #ExFor:FieldIncludeText.encoding
#     #ExFor:FieldIncludeText.lock_fields
#     #ExFor:FieldIncludeText.mime_type
#     #ExFor:FieldIncludeText.namespace_mappings
#     #ExFor:FieldIncludeText.source_full_name
#     #ExFor:FieldIncludeText.text_converter
#     #ExFor:FieldIncludeText.x_path
#     #ExFor:FieldIncludeText.xsl_transformation
#     #ExSummary:Shows how to create an INCLUDETEXT field, and set its properties.
#     [Test] #ExSkip
#     [Ignore("WORDSNET-17543")] #ExSkip
#     public void FieldIncludeText()
#
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Below are two ways to use INCLUDETEXT fields to display the contents of an XML file in the local file system.
#         # 1 -  Perform an XSL transformation on an XML document:
#         FieldIncludeText fieldIncludeText = CreateFieldIncludeText(builder, aeb.my_dir + "CD collection data.xml", false, "text/xml", "XML", "ISO-8859-1")
#         fieldIncludeText.xsl_transformation = aeb.my_dir + "CD collection XSL transformation.xsl"
#
#         builder.writeln()
#
#         # 2 -  Use an XPath to take specific elements from an XML document:
#         fieldIncludeText = CreateFieldIncludeText(builder, aeb.my_dir + "CD collection data.xml", false, "text/xml", "XML", "ISO-8859-1")
#         fieldIncludeText.namespace_mappings = "xmlns:n='myNamespace'"
#         fieldIncludeText.x_path = "/catalog/cd/title"
#
#         doc.save(aeb.artifacts_dir + "Field.includetext.docx")
#         TestFieldIncludeText(new Document(aeb.artifacts_dir + "Field.includetext.docx")) #ExSkip
#
#
#     # <summary>
#     # Use a document builder to insert an INCLUDETEXT field with custom properties.
#     # </summary>
#     public FieldIncludeText CreateFieldIncludeText(DocumentBuilder builder, string sourceFullName, bool lockFields, string mimeType, string textConverter, string encoding)
#
#         FieldIncludeText fieldIncludeText = (FieldIncludeText)builder.insert_field(FieldType.field_include_text, true)
#         fieldIncludeText.source_full_name = sourceFullName
#         fieldIncludeText.lock_fields = lockFields
#         fieldIncludeText.mime_type = mimeType
#         fieldIncludeText.text_converter = textConverter
#         fieldIncludeText.encoding = encoding
#
#         return fieldIncludeText
#
#     #ExEnd
#
#     private void TestFieldIncludeText(Document doc)
#
#         doc = DocumentHelper.save_open(doc)
#
#         FieldIncludeText fieldIncludeText = (FieldIncludeText)doc.range.fields[0]
#         self.assertEqual(aeb.my_dir + "CD collection data.xml", fieldIncludeText.source_full_name)
#         self.assertEqual(aeb.my_dir + "CD collection XSL transformation.xsl", fieldIncludeText.xsl_transformation)
#         self.assertFalse(fieldIncludeText.lock_fields)
#         self.assertEqual("text/xml", fieldIncludeText.mime_type)
#         self.assertEqual("XML", fieldIncludeText.text_converter)
#         self.assertEqual("ISO-8859-1", fieldIncludeText.encoding)
#         self.assertEqual(" INCLUDETEXT  \"" + aeb.my_dir.replace("\\", "\\\\") + "CD collection data.xml\" \\m text/xml \\c XML \\e ISO-8859-1 \\t \"" +
#                         aeb.my_dir.replace("\\", "\\\\") + "CD collection XSL transformation.xsl\"",
#             fieldIncludeText.get_field_code())
#         self.assertTrue(fieldIncludeText.result.starts_with("My CD Collection"))
#
#         XmlDocument cdCollectionData = new XmlDocument()
#         cdCollectionData.load_xml(File.read_all_text(aeb.my_dir + "CD collection data.xml"))
#         XmlNode catalogData = cdCollectionData.child_nodes[0]
#
#         XmlDocument cdCollectionXslTransformation = new XmlDocument()
#         cdCollectionXslTransformation.load_xml(File.read_all_text(aeb.my_dir + "CD collection XSL transformation.xsl"))
#
#         Table table = doc.first_section.body.tables[0]
#
#         XmlNamespaceManager manager = new XmlNamespaceManager(cdCollectionXslTransformation.name_table)
#         manager.add_namespace("xsl", "http:#www.w_3.org/1999/XSL/Transform")
#
#         for (int i = 0 i < table.rows.count i++)
#             for (int j = 0 j < table.rows[i].count j++)
#
#                 if (i == 0)
#
#                     # When on the first row from the input document's table, ensure that all table's cells match all XML element Names.
#                     for (int k = 0 k < table.rows.count - 1 k++)
#                         self.assertEqual(catalogData.child_nodes[k].child_nodes[j].name,
#                             table.rows[i].cells[j].get_text().replace(ControlChar.cell, string.empty).to_lower())
#
#                     # Also, make sure that the whole first row has the same color as the XSL transform.
#                     self.assertEqual(cdCollectionXslTransformation.select_nodes("#xsl:stylesheet/xsl:template/html/body/table/tr", manager)[0].attributes.get_named_item("bgcolor").value,
#                         ColorTranslator.to_html(table.rows[i].cells[j].cell_format.shading.background_pattern_color).to_lower())
#
#                 else
#
#                     # When on all other rows of the input document's table, ensure that cell contents match XML element Values.
#                     self.assertEqual(catalogData.child_nodes[i - 1].child_nodes[j].first_child.value,
#                         table.rows[i].cells[j].get_text().replace(ControlChar.cell, string.empty))
#                     self.assertEqual(Color.empty, table.rows[i].cells[j].cell_format.shading.background_pattern_color)
#
#
#                 self.assertEqual(
#                     double.parse(cdCollectionXslTransformation.select_nodes("#xsl:stylesheet/xsl:template/html/body/table", manager)[0].attributes.get_named_item("border").value) * 0.75,
#                     table.first_row.row_format.borders.bottom.line_width)
#
#
#         fieldIncludeText = (FieldIncludeText)doc.range.fields[1]
#         self.assertEqual(aeb.my_dir + "CD collection data.xml", fieldIncludeText.source_full_name)
#         Assert.null(fieldIncludeText.xsl_transformation)
#         self.assertFalse(fieldIncludeText.lock_fields)
#         self.assertEqual("text/xml", fieldIncludeText.mime_type)
#         self.assertEqual("XML", fieldIncludeText.text_converter)
#         self.assertEqual("ISO-8859-1", fieldIncludeText.encoding)
#         self.assertEqual(" INCLUDETEXT  \"" + aeb.my_dir.replace("\\", "\\\\") + "CD collection data.xml\" \\m text/xml \\c XML \\e ISO-8859-1 \\n xmlns:n='myNamespace' \\x /catalog/cd/title",
#             fieldIncludeText.get_field_code())
#
#         string expectedFieldResult = ""
#         for (int i = 0 i < catalogData.child_nodes.count i++)
#
#             expectedFieldResult += catalogData.child_nodes[i].child_nodes[0].child_nodes[0].value
#
#
#         self.assertEqual(expectedFieldResult, fieldIncludeText.result)
#
#
#     [Test]
#     [Ignore("WORDSNET-17545")]
#     public void FieldHyperlink()
#
#         #ExStart
#         #ExFor:FieldHyperlink
#         #ExFor:FieldHyperlink.address
#         #ExFor:FieldHyperlink.is_image_map
#         #ExFor:FieldHyperlink.open_in_new_window
#         #ExFor:FieldHyperlink.screen_tip
#         #ExFor:FieldHyperlink.sub_address
#         #ExFor:FieldHyperlink.target
#         #ExSummary:Shows how to use HYPERLINK fields to link to documents in the local file system.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         FieldHyperlink field = (FieldHyperlink)builder.insert_field(FieldType.field_hyperlink, true)
#
#         # When we click this HYPERLINK field in Microsoft Word,
#         # it will open the linked document and then place the cursor at the specified bookmark.
#         field.address = aeb.my_dir + "Bookmarks.docx"
#         field.sub_address = "MyBookmark3"
#         field.screen_tip = "Open " + field.address + " on bookmark " + field.sub_address + " in a new window"
#
#         builder.writeln()
#
#         # When we click this HYPERLINK field in Microsoft Word,
#         # it will open the linked document, and automatically scroll down to the specified iframe.
#         field = (FieldHyperlink)builder.insert_field(FieldType.field_hyperlink, true)
#         field.address = aeb.my_dir + "Iframes.html"
#         field.screen_tip = "Open " + field.address
#         field.target = "iframe_3"
#         field.open_in_new_window = true
#         field.is_image_map = false
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.hyperlink.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.hyperlink.docx")
#         field = (FieldHyperlink)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_hyperlink,
#             " HYPERLINK \"" + aeb.my_dir.replace("\\", "\\\\") + "Bookmarks.docx\" \\l \"MyBookmark3\" \\o \"Open " + aeb.my_dir + "Bookmarks.docx on bookmark MyBookmark3 in a new window\" ",
#             aeb.my_dir + "Bookmarks.docx - MyBookmark3", field)
#         self.assertEqual(aeb.my_dir + "Bookmarks.docx", field.address)
#         self.assertEqual("MyBookmark3", field.sub_address)
#         self.assertEqual("Open " + field.address.replace("\\", string.empty) + " on bookmark " + field.sub_address + " in a new window", field.screen_tip)
#
#         field = (FieldHyperlink)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_hyperlink, " HYPERLINK \"file:#" + aeb.my_dir.replace("\\", "\\\\").replace(" ", "%20") + "Iframes.html\" \\t \"iframe_3\" \\o \"Open " + aeb.my_dir.replace("\\", "\\\\") + "Iframes.html\" ",
#             aeb.my_dir + "Iframes.html", field)
#         self.assertEqual("file:#" + aeb.my_dir.replace(" ", "%20") + "Iframes.html", field.address)
#         self.assertEqual("Open " + aeb.my_dir + "Iframes.html", field.screen_tip)
#         self.assertEqual("iframe_3", field.target)
#         self.assertFalse(field.open_in_new_window)
#         self.assertFalse(field.is_image_map)
#
#
#     #ExStart
#     #ExFor:MergeFieldImageDimension
#     #ExFor:MergeFieldImageDimension.#ctor
#     #ExFor:MergeFieldImageDimension.#ctor(Double)
#     #ExFor:MergeFieldImageDimension.#ctor(Double,MergeFieldImageDimensionUnit)
#     #ExFor:MergeFieldImageDimension.unit
#     #ExFor:MergeFieldImageDimension.value
#     #ExFor:MergeFieldImageDimensionUnit
#     #ExFor:ImageFieldMergingArgs
#     #ExFor:ImageFieldMergingArgs.image_file_name
#     #ExFor:ImageFieldMergingArgs.image_width
#     #ExFor:ImageFieldMergingArgs.image_height
#     #ExSummary:Shows how to set the dimensions of images as MERGEFIELDS accepts them during a mail merge.
#     def test_merge_field_image_dimension(self) :
#
#         doc = aw.Document()
#
#         # Insert a MERGEFIELD that will accept images from a source during a mail merge. Use the field code to reference
#         # a column in the data source containing local system filenames of images we wish to use in the mail merge.
#         builder = aw.DocumentBuilder(doc)
#         FieldMergeField field = (FieldMergeField)builder.insert_field("MERGEFIELD Image:ImageColumn")
#
#         # The data source should have such a column named "ImageColumn".
#         self.assertEqual("Image:ImageColumn", field.field_name)
#
#         # Create a suitable data source.
#         DataTable dataTable = new DataTable("Images")
#         dataTable.columns.add(new DataColumn("ImageColumn"))
#         dataTable.rows.add(aeb.image_dir + "Logo.jpg")
#         dataTable.rows.add(aeb.image_dir + "Transparent background logo.png")
#         dataTable.rows.add(aeb.image_dir + "Enhanced Windows MetaFile.emf")
#
#         # Configure a callback to modify the sizes of images at merge time, then execute the mail merge.
#         doc.mail_merge.field_merging_callback = new MergedImageResizer(200, 200, MergeFieldImageDimensionUnit.point)
#         doc.mail_merge.execute(dataTable)
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.mergefield.image_dimension.docx")
#         TestMergeFieldImageDimension(doc) #ExSkip
#
#
#     # <summary>
#     # Sets the size of all mail merged images to one defined width and height.
#     # </summary>
#     private class MergedImageResizer : IFieldMergingCallback
#
#         public MergedImageResizer(double imageWidth, double imageHeight, MergeFieldImageDimensionUnit unit)
#
#             mImageWidth = imageWidth
#             mImageHeight = imageHeight
#             mUnit = unit
#
#
#         public void FieldMerging(FieldMergingArgs e)
#
#             throw new NotImplementedException()
#
#
#         public void ImageFieldMerging(ImageFieldMergingArgs args)
#
#             args.image_file_name = args.field_value.to_string()
#             args.image_width = new MergeFieldImageDimension(mImageWidth, mUnit)
#             args.image_height = new MergeFieldImageDimension(mImageHeight, mUnit)
#
#             self.assertEqual(mImageWidth, args.image_width.value)
#             self.assertEqual(mUnit, args.image_width.unit)
#             self.assertEqual(mImageHeight, args.image_height.value)
#             self.assertEqual(mUnit, args.image_height.unit)
#
#
#         private readonly double mImageWidth
#         private readonly double mImageHeight
#         private readonly MergeFieldImageDimensionUnit mUnit
#
#     #ExEnd
#
#     private void TestMergeFieldImageDimension(Document doc)
#
#         doc = DocumentHelper.save_open(doc)
#
#         self.assertEqual(0, doc.range.fields.count)
#         self.assertEqual(3, doc.get_child_nodes(NodeType.shape, true).count)
#
#         Shape shape = (Shape)doc.get_child(NodeType.shape, 0, true)
#
#         TestUtil.verify_image_in_shape(400, 400, ImageType.jpeg, shape)
#         self.assertEqual(200.0d, shape.width)
#         self.assertEqual(200.0d, shape.height)
#
#         shape = (Shape)doc.get_child(NodeType.shape, 1, true)
#
#         TestUtil.verify_image_in_shape(400, 400, ImageType.png, shape)
#         self.assertEqual(200.0d, shape.width)
#         self.assertEqual(200.0d, shape.height)
#
#         shape = (Shape)doc.get_child(NodeType.shape, 2, true)
#
#         TestUtil.verify_image_in_shape(534, 534, ImageType.emf, shape)
#         self.assertEqual(200.0d, shape.width)
#         self.assertEqual(200.0d, shape.height)
#
#
#     #ExStart
#     #ExFor:ImageFieldMergingArgs.image
#     #ExSummary:Shows how to use a callback to customize image merging logic.
#     def test_merge_field_images(self) :
#
#         doc = aw.Document()
#
#         # Insert a MERGEFIELD that will accept images from a source during a mail merge. Use the field code to reference
#         # a column in the data source which contains local system filenames of images we wish to use in the mail merge.
#         builder = aw.DocumentBuilder(doc)
#         FieldMergeField field = (FieldMergeField)builder.insert_field("MERGEFIELD Image:ImageColumn")
#
#         # In this case, the field expects the data source to have such a column named "ImageColumn".
#         self.assertEqual("Image:ImageColumn", field.field_name)
#
#         # Filenames can be lengthy, and if we can find a way to avoid storing them in the data source,
#         # we may considerably reduce its size.
#         # Create a data source that refers to images using short names.
#         DataTable dataTable = new DataTable("Images")
#         dataTable.columns.add(new DataColumn("ImageColumn"))
#         dataTable.rows.add("Dark logo")
#         dataTable.rows.add("Transparent logo")
#
#         # Assign a merging callback that contains all logic that processes those names,
#         # and then execute the mail merge.
#         doc.mail_merge.field_merging_callback = new ImageFilenameCallback()
#         doc.mail_merge.execute(dataTable)
#
#         doc.save(aeb.artifacts_dir + "Field.mergefield.images.docx")
#         TestMergeFieldImages(new Document(aeb.artifacts_dir + "Field.mergefield.images.docx")) #ExSkip
#
#
#     # <summary>
#     # Contains a dictionary that maps names of images to local system filenames that contain these images.
#     # If a mail merge data source uses one of the dictionary's names to refer to an image,
#     # this callback will pass the respective filename to the merge destination.
#     # </summary>
#     private class ImageFilenameCallback : IFieldMergingCallback
#
#         public ImageFilenameCallback()
#
#             mImageFilenames = new Dictionary<string, string>()
#             mImageFilenames.add("Dark logo", aeb.image_dir + "Logo.jpg")
#             mImageFilenames.add("Transparent logo", aeb.image_dir + "Transparent background logo.png")
#
#
#         void IFieldMergingCallback.field_merging(FieldMergingArgs args)
#
#             throw new NotImplementedException()
#
#
#         void IFieldMergingCallback.image_field_merging(ImageFieldMergingArgs args)
#
#             if (mImageFilenames.contains_key(args.field_value.to_string()))
#
#                 #if NET462 || JAVA
#                 args.image = Image.from_file(mImageFilenames[args.field_value.to_string()])
#                 #elif NETCOREAPP2_1
#                 args.image = SKBitmap.decode(mImageFilenames[args.field_value.to_string()])
#                 args.image_file_name = mImageFilenames[args.field_value.to_string()]
#                 #endif
#
#
#             Assert.not_null(args.image)
#
#
#         private readonly Dictionary<string, string> mImageFilenames
#
#     #ExEnd
#
#     private void TestMergeFieldImages(Document doc)
#
#         doc = DocumentHelper.save_open(doc)
#
#         self.assertEqual(0, doc.range.fields.count)
#         self.assertEqual(2, doc.get_child_nodes(NodeType.shape, true).count)
#
#         Shape shape = (Shape)doc.get_child(NodeType.shape, 0, true)
#
#         TestUtil.verify_image_in_shape(400, 400, ImageType.jpeg, shape)
#         self.assertEqual(300.0d, shape.width)
#         self.assertEqual(300.0d, shape.height)
#
#         shape = (Shape)doc.get_child(NodeType.shape, 1, true)
#
#         TestUtil.verify_image_in_shape(400, 400, ImageType.png, shape)
#         self.assertEqual(300.0d, shape.width, 1)
#         self.assertEqual(300.0d, shape.height, 1)
#
#
#     [Test]
#     [Ignore("WORDSNET-17524")]
#     public void FieldIndexFilter()
#
#         #ExStart
#         #ExFor:FieldIndex
#         #ExFor:FieldIndex.bookmark_name
#         #ExFor:FieldIndex.entry_type
#         #ExFor:FieldXE
#         #ExFor:FieldXE.entry_type
#         #ExFor:FieldXE.text
#         #ExSummary:Shows how to create an INDEX field, and then use XE fields to populate it with entries.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Create an INDEX field which will display an entry for each XE field found in the document.
#         # Each entry will display the XE field's Text property value on the left side
#         # and the page containing the XE field on the right.
#         # If the XE fields have the same value in their "Text" property,
#         # the INDEX field will group them into one entry.
#         FieldIndex index = (FieldIndex)builder.insert_field(FieldType.field_index, true)
#
#         # Configure the INDEX field only to display XE fields that are within the bounds
#         # of a bookmark named "MainBookmark", and whose "EntryType" properties have a value of "A".
#         # For both INDEX and XE fields, the "EntryType" property only uses the first character of its string value.
#         index.bookmark_name = "MainBookmark"
#         index.entry_type = "A"
#
#         self.assertEqual(" INDEX  \\b MainBookmark \\f A", index.get_field_code())
#
#         # On a new page, start the bookmark with a name that matches the value
#         # of the INDEX field's "BookmarkName" property.
#         builder.insert_break(BreakType.page_break)
#         builder.start_bookmark("MainBookmark")
#
#         # The INDEX field will pick up this entry because it is inside the bookmark,
#         # and its entry type also matches the INDEX field's entry type.
#         FieldXE indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "Index entry 1"
#         indexEntry.entry_type = "A"
#
#         self.assertEqual(" XE  \"Index entry 1\" \\f A", indexEntry.get_field_code())
#
#         # Insert an XE field that will not appear in the INDEX because the entry types do not match.
#         builder.insert_break(BreakType.page_break)
#         indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "Index entry 2"
#         indexEntry.entry_type = "B"
#
#         # End the bookmark and insert an XE field afterwards.
#         # It is of the same type as the INDEX field, but will not appear
#         # since it is outside the bookmark's boundaries.
#         builder.end_bookmark("MainBookmark")
#         builder.insert_break(BreakType.page_break)
#         indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "Index entry 3"
#         indexEntry.entry_type = "A"
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.index.xe.filtering.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.index.xe.filtering.docx")
#         index = (FieldIndex)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_index, " INDEX  \\b MainBookmark \\f A", "Index entry 1, 2\r", index)
#         self.assertEqual("MainBookmark", index.bookmark_name)
#         self.assertEqual("A", index.entry_type)
#
#         indexEntry = (FieldXE)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " XE  \"Index entry 1\" \\f A", string.empty, indexEntry)
#         self.assertEqual("Index entry 1", indexEntry.text)
#         self.assertEqual("A", indexEntry.entry_type)
#
#         indexEntry = (FieldXE)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " XE  \"Index entry 2\" \\f B", string.empty, indexEntry)
#         self.assertEqual("Index entry 2", indexEntry.text)
#         self.assertEqual("B", indexEntry.entry_type)
#
#         indexEntry = (FieldXE)doc.range.fields[3]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " XE  \"Index entry 3\" \\f A", string.empty, indexEntry)
#         self.assertEqual("Index entry 3", indexEntry.text)
#         self.assertEqual("A", indexEntry.entry_type)
#
#
#     [Test]
#     [Ignore("WORDSNET-17524")]
#     public void FieldIndexFormatting()
#
#         #ExStart
#         #ExFor:FieldIndex
#         #ExFor:FieldIndex.heading
#         #ExFor:FieldIndex.number_of_columns
#         #ExFor:FieldIndex.language_id
#         #ExFor:FieldIndex.letter_range
#         #ExFor:FieldXE
#         #ExFor:FieldXE.is_bold
#         #ExFor:FieldXE.is_italic
#         #ExFor:FieldXE.text
#         #ExSummary:Shows how to populate an INDEX field with entries using XE fields, and also modify its appearance.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Create an INDEX field which will display an entry for each XE field found in the document.
#         # Each entry will display the XE field's Text property value on the left side,
#         # and the number of the page that contains the XE field on the right.
#         # If the XE fields have the same value in their "Text" property,
#         # the INDEX field will group them into one entry.
#         FieldIndex index = (FieldIndex)builder.insert_field(FieldType.field_index, true)
#         index.language_id = "1033"
#
#         # Setting this property's value to "A" will group all the entries by their first letter,
#         # and place that letter in uppercase above each group.
#         index.heading = "A"
#
#         # Set the table created by the INDEX field to span over 2 columns.
#         index.number_of_columns = "2"
#
#         # Set any entries with starting letters outside the "a-c" character range to be omitted.
#         index.letter_range = "a-c"
#
#         self.assertEqual(" INDEX  \\z 1033 \\h A \\c 2 \\p a-c", index.get_field_code())
#
#         # These next two XE fields will show up under the "A" heading,
#         # with their respective text stylings also applied to their page numbers.
#         builder.insert_break(BreakType.page_break)
#         FieldXE indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "Apple"
#         indexEntry.is_italic = true
#
#         self.assertEqual(" XE  Apple \\i", indexEntry.get_field_code())
#
#         builder.insert_break(BreakType.page_break)
#         indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "Apricot"
#         indexEntry.is_bold = true
#
#         self.assertEqual(" XE  Apricot \\b", indexEntry.get_field_code())
#
#         # Both the next two XE fields will be under a "B" and "C" heading in the INDEX fields table of contents.
#         builder.insert_break(BreakType.page_break)
#         indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "Banana"
#
#         builder.insert_break(BreakType.page_break)
#         indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "Cherry"
#
#         # INDEX fields sort all entries alphabetically, so this entry will show up under "A" with the other two.
#         builder.insert_break(BreakType.page_break)
#         indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "Avocado"
#
#         # This entry will not appear because it starts with the letter "D",
#         # which is outside the "a-c" character range that the INDEX field's LetterRange property defines.
#         builder.insert_break(BreakType.page_break)
#         indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "Durian"
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.index.xe.formatting.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.index.xe.formatting.docx")
#         index = (FieldIndex)doc.range.fields[0]
#
#         self.assertEqual("1033", index.language_id)
#         self.assertEqual("A", index.heading)
#         self.assertEqual("2", index.number_of_columns)
#         self.assertEqual("a-c", index.letter_range)
#         self.assertEqual(" INDEX  \\z 1033 \\h A \\c 2 \\p a-c", index.get_field_code())
#         self.assertEqual("\fA\r" +
#                         "Apple, 2\r" +
#                         "Apricot, 3\r" +
#                         "Avocado, 6\r" +
#                         "B\r" +
#                         "Banana, 4\r" +
#                         "C\r" +
#                         "Cherry, 5\r\f", index.result)
#
#         indexEntry = (FieldXE)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " XE  Apple \\i", string.empty, indexEntry)
#         self.assertEqual("Apple", indexEntry.text)
#         self.assertFalse(indexEntry.is_bold)
#         self.assertTrue(indexEntry.is_italic)
#
#         indexEntry = (FieldXE)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " XE  Apricot \\b", string.empty, indexEntry)
#         self.assertEqual("Apricot", indexEntry.text)
#         self.assertTrue(indexEntry.is_bold)
#         self.assertFalse(indexEntry.is_italic)
#
#         indexEntry = (FieldXE)doc.range.fields[3]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " XE  Banana", string.empty, indexEntry)
#         self.assertEqual("Banana", indexEntry.text)
#         self.assertFalse(indexEntry.is_bold)
#         self.assertFalse(indexEntry.is_italic)
#
#         indexEntry = (FieldXE)doc.range.fields[4]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " XE  Cherry", string.empty, indexEntry)
#         self.assertEqual("Cherry", indexEntry.text)
#         self.assertFalse(indexEntry.is_bold)
#         self.assertFalse(indexEntry.is_italic)
#
#         indexEntry = (FieldXE)doc.range.fields[5]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " XE  Avocado", string.empty, indexEntry)
#         self.assertEqual("Avocado", indexEntry.text)
#         self.assertFalse(indexEntry.is_bold)
#         self.assertFalse(indexEntry.is_italic)
#
#         indexEntry = (FieldXE)doc.range.fields[6]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " XE  Durian", string.empty, indexEntry)
#         self.assertEqual("Durian", indexEntry.text)
#         self.assertFalse(indexEntry.is_bold)
#         self.assertFalse(indexEntry.is_italic)
#
#
#     [Test]
#     [Ignore("WORDSNET-17524")]
#     public void FieldIndexSequence()
#
#         #ExStart
#         #ExFor:FieldIndex.has_sequence_name
#         #ExFor:FieldIndex.sequence_name
#         #ExFor:FieldIndex.sequence_separator
#         #ExSummary:Shows how to split a document into portions by combining INDEX and SEQ fields.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Create an INDEX field which will display an entry for each XE field found in the document.
#         # Each entry will display the XE field's Text property value on the left side,
#         # and the number of the page that contains the XE field on the right.
#         # If the XE fields have the same value in their "Text" property,
#         # the INDEX field will group them into one entry.
#         FieldIndex index = (FieldIndex)builder.insert_field(FieldType.field_index, true)
#
#         # In the SequenceName property, name a SEQ field sequence. Each entry of this INDEX field will now also display
#         # the number that the sequence count is on at the XE field location that created this entry.
#         index.sequence_name = "MySequence"
#
#         # Set text that will around the sequence and page numbers to explain their meaning to the user.
#         # An entry created with this configuration will display something like "MySequence at 1 on page 1" at its page number.
#         # PageNumberSeparator and SequenceSeparator cannot be longer than 15 characters.
#         index.page_number_separator = "\tMySequence at "
#         index.sequence_separator = " on page "
#         self.assertTrue(index.has_sequence_name)
#
#         self.assertEqual(" INDEX  \\s MySequence \\e \"\tMySequence at \" \\d \" on page \"", index.get_field_code())
#
#         # SEQ fields display a count that increments at each SEQ field.
#         # These fields also maintain separate counts for each unique named sequence
#         # identified by the SEQ field's "SequenceIdentifier" property.
#         # Insert a SEQ field which moves the "MySequence" sequence to 1.
#         # This field no different from normal document text. It will not appear on an INDEX field's table of contents.
#         builder.insert_break(BreakType.page_break)
#         FieldSeq sequenceField = (FieldSeq)builder.insert_field(FieldType.field_sequence, true)
#         sequenceField.sequence_identifier = "MySequence"
#
#         self.assertEqual(" SEQ  MySequence", sequenceField.get_field_code())
#
#         # Insert an XE field which will create an entry in the INDEX field.
#         # Since "MySequence" is at 1 and this XE field is on page 2, along with the custom separators we defined above,
#         # this field's INDEX entry will display "Cat" on the left side, and "MySequence at 1 on page 2" on the right.
#         FieldXE indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "Cat"
#
#         self.assertEqual(" XE  Cat", indexEntry.get_field_code())
#
#         # Insert a page break and use SEQ fields to advance "MySequence" to 3.
#         builder.insert_break(BreakType.page_break)
#         sequenceField = (FieldSeq)builder.insert_field(FieldType.field_sequence, true)
#         sequenceField.sequence_identifier = "MySequence"
#         sequenceField = (FieldSeq)builder.insert_field(FieldType.field_sequence, true)
#         sequenceField.sequence_identifier = "MySequence"
#
#         # Insert an XE field with the same Text property as the one above.
#         # The INDEX entry will group XE fields with matching values in the "Text" property
#         # into one entry as opposed to making an entry for each XE field.
#         # Since we are on page 2 with "MySequence" at 3, ", 3 on page 3" will be appended to the same INDEX entry as above.
#         # The page number portion of that INDEX entry will now display "MySequence at 1 on page 2, 3 on page 3".
#         indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "Cat"
#
#         # Insert an XE field with a new and unique Text property value.
#         # This will add a new entry, with MySequence at 3 on page 4.
#         builder.insert_break(BreakType.page_break)
#         indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "Dog"
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.index.xe.sequence.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.index.xe.sequence.docx")
#         index = (FieldIndex)doc.range.fields[0]
#
#         self.assertEqual("MySequence", index.sequence_name)
#         self.assertEqual("\tMySequence at ", index.page_number_separator)
#         self.assertEqual(" on page ", index.sequence_separator)
#         self.assertTrue(index.has_sequence_name)
#         self.assertEqual(" INDEX  \\s MySequence \\e \"\tMySequence at \" \\d \" on page \"", index.get_field_code())
#         self.assertEqual("Cat\tMySequence at 1 on page 2, 3 on page 3\r" +
#                         "Dog\tMySequence at 3 on page 4\r", index.result)
#
#         self.assertEqual(3, doc.range.fields.where(f => f.type == FieldType.field_sequence).count())
#
#
#     [Test]
#     [Ignore("WORDSNET-17524")]
#     public void FieldIndexPageNumberSeparator()
#
#         #ExStart
#         #ExFor:FieldIndex.has_page_number_separator
#         #ExFor:FieldIndex.page_number_separator
#         #ExFor:FieldIndex.page_number_list_separator
#         #ExSummary:Shows how to edit the page number separator in an INDEX field.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Create an INDEX field which will display an entry for each XE field found in the document.
#         # Each entry will display the XE field's Text property value on the left side,
#         # and the number of the page that contains the XE field on the right.
#         # The INDEX entry will group XE fields with matching values in the "Text" property
#         # into one entry as opposed to making an entry for each XE field.
#         FieldIndex index = (FieldIndex)builder.insert_field(FieldType.field_index, true)
#
#         # If our INDEX field has an entry for a group of XE fields,
#         # this entry will display the number of each page that contains an XE field that belongs to this group.
#         # We can set custom separators to customize the appearance of these page numbers.
#         index.page_number_separator = ", on page(s) "
#         index.page_number_list_separator = " & "
#
#         self.assertEqual(" INDEX  \\e \", on page(s) \" \\l \" & \"", index.get_field_code())
#         self.assertTrue(index.has_page_number_separator)
#
#         # After we insert these XE fields, the INDEX field will display "First entry, on page(s) 2 & 3 & 4".
#         builder.insert_break(BreakType.page_break)
#         FieldXE indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "First entry"
#
#         self.assertEqual(" XE  \"First entry\"", indexEntry.get_field_code())
#
#         builder.insert_break(BreakType.page_break)
#         indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "First entry"
#
#         builder.insert_break(BreakType.page_break)
#         indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "First entry"
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.index.xe.page_number_list.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.index.xe.page_number_list.docx")
#         index = (FieldIndex)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_index, " INDEX  \\e \", on page(s) \" \\l \" & \"", "First entry, on page(s) 2 & 3 & 4\r", index)
#         self.assertEqual(", on page(s) ", index.page_number_separator)
#         self.assertEqual(" & ", index.page_number_list_separator)
#         self.assertTrue(index.has_page_number_separator)
#
#
#     [Test]
#     [Ignore("WORDSNET-17524")]
#     public void FieldIndexPageRangeBookmark()
#
#         #ExStart
#         #ExFor:FieldIndex.page_range_separator
#         #ExFor:FieldXE.has_page_range_bookmark_name
#         #ExFor:FieldXE.page_range_bookmark_name
#         #ExSummary:Shows how to specify a bookmark's spanned pages as a page range for an INDEX field entry.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Create an INDEX field which will display an entry for each XE field found in the document.
#         # Each entry will display the XE field's Text property value on the left side,
#         # and the number of the page that contains the XE field on the right.
#         # The INDEX entry will collect all XE fields with matching values in the "Text" property
#         # into one entry as opposed to making an entry for each XE field.
#         FieldIndex index = (FieldIndex)builder.insert_field(FieldType.field_index, true)
#
#         # For INDEX entries that display page ranges, we can specify a separator string
#         # which will appear between the number of the first page, and the number of the last.
#         index.page_number_separator = ", on page(s) "
#         index.page_range_separator = " to "
#
#         self.assertEqual(" INDEX  \\e \", on page(s) \" \\g \" to \"", index.get_field_code())
#
#         builder.insert_break(BreakType.page_break)
#         FieldXE indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "My entry"
#
#         # If an XE field names a bookmark using the PageRangeBookmarkName property,
#         # its INDEX entry will show the range of pages that the bookmark spans
#         # instead of the number of the page that contains the XE field.
#         indexEntry.page_range_bookmark_name = "MyBookmark"
#
#         self.assertEqual(" XE  \"My entry\" \\r MyBookmark", indexEntry.get_field_code())
#         self.assertTrue(indexEntry.has_page_range_bookmark_name)
#
#         # Insert a bookmark that starts on page 3 and ends on page 5.
#         # The INDEX entry for the XE field that references this bookmark will display this page range.
#         # In our table, the INDEX entry will display "My entry, on page(s) 3 to 5".
#         builder.insert_break(BreakType.page_break)
#         builder.start_bookmark("MyBookmark")
#         builder.write("Start of MyBookmark")
#         builder.insert_break(BreakType.page_break)
#         builder.insert_break(BreakType.page_break)
#         builder.write("End of MyBookmark")
#         builder.end_bookmark("MyBookmark")
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.index.xe.page_range_bookmark.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.index.xe.page_range_bookmark.docx")
#         index = (FieldIndex)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_index, " INDEX  \\e \", on page(s) \" \\g \" to \"", "My entry, on page(s) 3 to 5\r", index)
#         self.assertEqual(", on page(s) ", index.page_number_separator)
#         self.assertEqual(" to ", index.page_range_separator)
#
#         indexEntry = (FieldXE)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " XE  \"My entry\" \\r MyBookmark", string.empty, indexEntry)
#         self.assertEqual("My entry", indexEntry.text)
#         self.assertEqual("MyBookmark", indexEntry.page_range_bookmark_name)
#         self.assertTrue(indexEntry.has_page_range_bookmark_name)
#
#
#     [Test]
#     [Ignore("WORDSNET-17524")]
#     public void FieldIndexCrossReferenceSeparator()
#
#         #ExStart
#         #ExFor:FieldIndex.cross_reference_separator
#         #ExFor:FieldXE.page_number_replacement
#         #ExSummary:Shows how to define cross references in an INDEX field.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Create an INDEX field which will display an entry for each XE field found in the document.
#         # Each entry will display the XE field's Text property value on the left side,
#         # and the number of the page that contains the XE field on the right.
#         # The INDEX entry will collect all XE fields with matching values in the "Text" property
#         # into one entry as opposed to making an entry for each XE field.
#         FieldIndex index = (FieldIndex)builder.insert_field(FieldType.field_index, true)
#
#         # We can configure an XE field to get its INDEX entry to display a string instead of a page number.
#         # First, for entries that substitute a page number with a string,
#         # specify a custom separator between the XE field's Text property value and the string.
#         index.cross_reference_separator = ", see: "
#
#         self.assertEqual(" INDEX  \\k \", see: \"", index.get_field_code())
#
#         # Insert an XE field, which creates a regular INDEX entry which displays this field's page number,
#         # and does not invoke the CrossReferenceSeparator value.
#         # The entry for this XE field will display "Apple, 2".
#         builder.insert_break(BreakType.page_break)
#         FieldXE indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "Apple"
#
#         self.assertEqual(" XE  Apple", indexEntry.get_field_code())
#
#         # Insert another XE field on page 3 and set a value for the PageNumberReplacement property.
#         # This value will show up instead of the number of the page that this field is on,
#         # and the INDEX field's CrossReferenceSeparator value will appear in front of it.
#         # The entry for this XE field will display "Banana, see: Tropical fruit".
#         builder.insert_break(BreakType.page_break)
#         indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "Banana"
#         indexEntry.page_number_replacement = "Tropical fruit"
#
#         self.assertEqual(" XE  Banana \\t \"Tropical fruit\"", indexEntry.get_field_code())
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.index.xe.cross_reference_separator.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.index.xe.cross_reference_separator.docx")
#         index = (FieldIndex)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " INDEX  \\k \", see: \"",
#             "Apple, 2\r" +
#             "Banana, see: Tropical fruit\r", index)
#         self.assertEqual(", see: ", index.cross_reference_separator)
#
#         indexEntry = (FieldXE)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " XE  Apple", string.empty, indexEntry)
#         self.assertEqual("Apple", indexEntry.text)
#         Assert.null(indexEntry.page_number_replacement)
#
#         indexEntry = (FieldXE)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " XE  Banana \\t \"Tropical fruit\"", string.empty, indexEntry)
#         self.assertEqual("Banana", indexEntry.text)
#         self.assertEqual("Tropical fruit", indexEntry.page_number_replacement)
#
#
#     [TestCase(true)]
#     [TestCase(false)]
#     [Ignore("WORDSNET-17524")]
#     public void FieldIndexSubheading(bool runSubentriesOnTheSameLine)
#
#         #ExStart
#         #ExFor:FieldIndex.run_subentries_on_same_line
#         #ExSummary:Shows how to work with subentries in an INDEX field.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Create an INDEX field which will display an entry for each XE field found in the document.
#         # Each entry will display the XE field's Text property value on the left side,
#         # and the number of the page that contains the XE field on the right.
#         # The INDEX entry will collect all XE fields with matching values in the "Text" property
#         # into one entry as opposed to making an entry for each XE field.
#         FieldIndex index = (FieldIndex)builder.insert_field(FieldType.field_index, true)
#         index.page_number_separator = ", see page "
#         index.heading = "A"
#
#         # XE fields that have a Text property whose value becomes the heading of the INDEX entry.
#         # If this value contains two string segments split by a colon (the INDEX entry will treat :) delimiter,
#         # the first segment is heading, and the second segment will become the subheading.
#         # The INDEX field first groups entries alphabetically, then, if there are multiple XE fields with the same
#         # headings, the INDEX field will further subgroup them by the values of these headings.
#         # There can be multiple subgrouping layers, depending on how many times
#         # the Text properties of XE fields get segmented like this.
#         # By default, an INDEX field entry group will create a new line for every subheading within this group.
#         # We can set the RunSubentriesOnSameLine flag to true to keep the heading,
#         # and every subheading for the group on one line instead, which will make the INDEX field more compact.
#         index.run_subentries_on_same_line = runSubentriesOnTheSameLine
#
#         if (runSubentriesOnTheSameLine)
#             self.assertEqual(" INDEX  \\e \", see page \" \\h A \\r", index.get_field_code())
#         else
#             self.assertEqual(" INDEX  \\e \", see page \" \\h A", index.get_field_code())
#
#         # Insert two XE fields, each on a new page, and with the same heading named "Heading 1",
#         # which the INDEX field will use to group them.
#         # If RunSubentriesOnSameLine is false, then the INDEX table will create three lines:
#         # one line for the grouping heading "Heading 1", and one more line for each subheading.
#         # If RunSubentriesOnSameLine is true, then the INDEX table will create a one-line
#         # entry that encompasses the heading and every subheading.
#         builder.insert_break(BreakType.page_break)
#         FieldXE indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "Heading 1:Subheading 1"
#
#         self.assertEqual(" XE  \"Heading 1:Subheading 1\"", indexEntry.get_field_code())
#
#         builder.insert_break(BreakType.page_break)
#         indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "Heading 1:Subheading 2"
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + $"Field.index.xe.subheading.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + $"Field.index.xe.subheading.docx")
#         index = (FieldIndex)doc.range.fields[0]
#
#         if (runSubentriesOnTheSameLine)
#
#             TestUtil.verify_field(FieldType.field_index, " INDEX  \\r \\e \", see page \" \\h A",
#                 "H\r" +
#                 "Heading 1: Subheading 1, see page 2 Subheading 2, see page 3\r", index)
#             self.assertTrue(index.run_subentries_on_same_line)
#
#         else
#
#             TestUtil.verify_field(FieldType.field_index, " INDEX  \\e \", see page \" \\h A",
#                 "H\r" +
#                 "Heading 1\r" +
#                 "Subheading 1, see page 2\r" +
#                 "Subheading 2, see page 3\r", index)
#             self.assertFalse(index.run_subentries_on_same_line)
#
#
#         indexEntry = (FieldXE)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " XE  \"Heading 1:Subheading 1\"", string.empty, indexEntry)
#         self.assertEqual("Heading 1:Subheading 1", indexEntry.text)
#
#         indexEntry = (FieldXE)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " XE  \"Heading 1:Subheading 2\"", string.empty, indexEntry)
#         self.assertEqual("Heading 1:Subheading 2", indexEntry.text)
#
#
#     [TestCase(true)]
#     [TestCase(false)]
#     [Ignore("WORDSNET-17524")]
#     public void FieldIndexYomi(bool sortEntriesUsingYomi)
#
#         #ExStart
#         #ExFor:FieldIndex.use_yomi
#         #ExFor:FieldXE.yomi
#         #ExSummary:Shows how to sort INDEX field entries phonetically.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Create an INDEX field which will display an entry for each XE field found in the document.
#         # Each entry will display the XE field's Text property value on the left side,
#         # and the number of the page that contains the XE field on the right.
#         # The INDEX entry will collect all XE fields with matching values in the "Text" property
#         # into one entry as opposed to making an entry for each XE field.
#         FieldIndex index = (FieldIndex)builder.insert_field(FieldType.field_index, true)
#
#         # The INDEX table automatically sorts its entries by the values of their Text properties in alphabetic order.
#         # Set the INDEX table to sort entries phonetically using Hiragana instead.
#         index.use_yomi = sortEntriesUsingYomi
#
#         if (sortEntriesUsingYomi)
#             self.assertEqual(" INDEX  \\y", index.get_field_code())
#         else
#             self.assertEqual(" INDEX ", index.get_field_code())
#
#         # Insert 4 XE fields, which would show up as entries in the INDEX field's table of contents.
#         # The "Text" property may contain a word's spelling in Kanji, whose pronunciation may be ambiguous,
#         # while the "Yomi" version of the word will spell exactly how it is pronounced using Hiragana.
#         # If we set our INDEX field to use Yomi, it will sort these entries
#         # by the value of their Yomi properties, instead of their Text values.
#         builder.insert_break(BreakType.page_break)
#         FieldXE indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "愛子"
#         indexEntry.yomi = "あ"
#
#         self.assertEqual(" XE  愛子 \\y あ", indexEntry.get_field_code())
#
#         builder.insert_break(BreakType.page_break)
#         indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "明美"
#         indexEntry.yomi = "あ"
#
#         builder.insert_break(BreakType.page_break)
#         indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "恵美"
#         indexEntry.yomi = "え"
#
#         builder.insert_break(BreakType.page_break)
#         indexEntry = (FieldXE)builder.insert_field(FieldType.field_index_entry, true)
#         indexEntry.text = "愛美"
#         indexEntry.yomi = "え"
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.index.xe.yomi.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.index.xe.yomi.docx")
#         index = (FieldIndex)doc.range.fields[0]
#
#         if (sortEntriesUsingYomi)
#
#             self.assertTrue(index.use_yomi)
#             self.assertEqual(" INDEX  \\y", index.get_field_code())
#             self.assertEqual("愛子, 2\r" +
#                             "明美, 3\r" +
#                             "恵美, 4\r" +
#                             "愛美, 5\r", index.result)
#
#         else
#
#             self.assertFalse(index.use_yomi)
#             self.assertEqual(" INDEX ", index.get_field_code())
#             self.assertEqual("恵美, 4\r" +
#                             "愛子, 2\r" +
#                             "愛美, 5\r" +
#                             "明美, 3\r", index.result)
#
#
#         indexEntry = (FieldXE)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " XE  愛子 \\y あ", string.empty, indexEntry)
#         self.assertEqual("愛子", indexEntry.text)
#         self.assertEqual("あ", indexEntry.yomi)
#
#         indexEntry = (FieldXE)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " XE  明美 \\y あ", string.empty, indexEntry)
#         self.assertEqual("明美", indexEntry.text)
#         self.assertEqual("あ", indexEntry.yomi)
#
#         indexEntry = (FieldXE)doc.range.fields[3]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " XE  恵美 \\y え", string.empty, indexEntry)
#         self.assertEqual("恵美", indexEntry.text)
#         self.assertEqual("え", indexEntry.yomi)
#
#         indexEntry = (FieldXE)doc.range.fields[4]
#
#         TestUtil.verify_field(FieldType.field_index_entry, " XE  愛美 \\y え", string.empty, indexEntry)
#         self.assertEqual("愛美", indexEntry.text)
#         self.assertEqual("え", indexEntry.yomi)
#
#
#     def test_field_barcode(self) :
#
#         #ExStart
#         #ExFor:FieldBarcode
#         #ExFor:FieldBarcode.facing_identification_mark
#         #ExFor:FieldBarcode.is_bookmark
#         #ExFor:FieldBarcode.is_us_postal_address
#         #ExFor:FieldBarcode.postal_address
#         #ExSummary:Shows how to use the BARCODE field to display U.s. ZIP codes in the form of a barcode.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         builder.writeln()
#
#         # Below are two ways of using BARCODE fields to display custom values as barcodes.
#         # 1 -  Store the value that the barcode will display in the PostalAddress property:
#         FieldBarcode field = (FieldBarcode)builder.insert_field(FieldType.field_barcode, true)
#
#         # This value needs to be a valid ZIP code.
#         field.postal_address = "96801"
#         field.is_us_postal_address = true
#         field.facing_identification_mark = "C"
#
#         self.assertEqual(" BARCODE  96801 \\u \\f C", field.get_field_code())
#
#         builder.insert_break(BreakType.line_break)
#
#         # 2 -  Reference a bookmark that stores the value that this barcode will display:
#         field = (FieldBarcode)builder.insert_field(FieldType.field_barcode, true)
#         field.postal_address = "BarcodeBookmark"
#         field.is_bookmark = true
#
#         self.assertEqual(" BARCODE  BarcodeBookmark \\b", field.get_field_code())
#
#         # The bookmark that the BARCODE field references in its PostalAddress property
#         # need to contain nothing besides the valid ZIP code.
#         builder.insert_break(BreakType.page_break)
#         builder.start_bookmark("BarcodeBookmark")
#         builder.writeln("968877")
#         builder.end_bookmark("BarcodeBookmark")
#
#         doc.save(aeb.artifacts_dir + "Field.barcode.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.barcode.docx")
#
#         self.assertEqual(0, doc.get_child_nodes(NodeType.shape, true).count)
#
#         field = (FieldBarcode)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_barcode, " BARCODE  96801 \\u \\f C", string.empty, field)
#         self.assertEqual("C", field.facing_identification_mark)
#         self.assertEqual("96801", field.postal_address)
#         self.assertTrue(field.is_us_postal_address)
#
#         field = (FieldBarcode)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_barcode, " BARCODE  BarcodeBookmark \\b", string.empty, field)
#         self.assertEqual("BarcodeBookmark", field.postal_address)
#         self.assertTrue(field.is_bookmark)
#
#
#     def test_field_display_barcode(self) :
#
#         #ExStart
#         #ExFor:FieldDisplayBarcode
#         #ExFor:FieldDisplayBarcode.add_start_stop_char
#         #ExFor:FieldDisplayBarcode.background_color
#         #ExFor:FieldDisplayBarcode.barcode_type
#         #ExFor:FieldDisplayBarcode.barcode_value
#         #ExFor:FieldDisplayBarcode.case_code_style
#         #ExFor:FieldDisplayBarcode.display_text
#         #ExFor:FieldDisplayBarcode.error_correction_level
#         #ExFor:FieldDisplayBarcode.fix_check_digit
#         #ExFor:FieldDisplayBarcode.foreground_color
#         #ExFor:FieldDisplayBarcode.pos_code_style
#         #ExFor:FieldDisplayBarcode.scaling_factor
#         #ExFor:FieldDisplayBarcode.symbol_height
#         #ExFor:FieldDisplayBarcode.symbol_rotation
#         #ExSummary:Shows how to insert a DISPLAYBARCODE field, and set its properties.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         FieldDisplayBarcode field = (FieldDisplayBarcode)builder.insert_field(FieldType.field_display_barcode, true)
#
#         # Below are four types of barcodes, decorated in various ways, that the DISPLAYBARCODE field can display.
#         # 1 -  QR code with custom colors:
#         field.barcode_type = "QR"
#         field.barcode_value = "ABC123"
#         field.background_color = "0xF8BD69"
#         field.foreground_color = "0xB5413B"
#         field.error_correction_level = "3"
#         field.scaling_factor = "250"
#         field.symbol_height = "1000"
#         field.symbol_rotation = "0"
#
#         self.assertEqual(" DISPLAYBARCODE  ABC123 QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0", field.get_field_code())
#         builder.writeln()
#
#         # 2 -  EAN13 barcode, with the digits displayed below the bars:
#         field = (FieldDisplayBarcode)builder.insert_field(FieldType.field_display_barcode, true)
#         field.barcode_type = "EAN13"
#         field.barcode_value = "501234567890"
#         field.display_text = true
#         field.pos_code_style = "CASE"
#         field.fix_check_digit = true
#
#         self.assertEqual(" DISPLAYBARCODE  501234567890 EAN13 \\t \\p CASE \\x", field.get_field_code())
#         builder.writeln()
#
#         # 3 -  CODE39 barcode:
#         field = (FieldDisplayBarcode)builder.insert_field(FieldType.field_display_barcode, true)
#         field.barcode_type = "CODE39"
#         field.barcode_value = "12345ABCDE"
#         field.add_start_stop_char = true
#
#         self.assertEqual(" DISPLAYBARCODE  12345ABCDE CODE39 \\d", field.get_field_code())
#         builder.writeln()
#
#         # 4 -  ITF4 barcode, with a specified case code:
#         field = (FieldDisplayBarcode)builder.insert_field(FieldType.field_display_barcode, true)
#         field.barcode_type = "ITF14"
#         field.barcode_value = "09312345678907"
#         field.case_code_style = "STD"
#
#         self.assertEqual(" DISPLAYBARCODE  09312345678907 ITF14 \\c STD", field.get_field_code())
#
#         doc.save(aeb.artifacts_dir + "Field.displaybarcode.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.displaybarcode.docx")
#
#         self.assertEqual(0, doc.get_child_nodes(NodeType.shape, true).count)
#
#         field = (FieldDisplayBarcode)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_display_barcode, " DISPLAYBARCODE  ABC123 QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0", string.empty, field)
#         self.assertEqual("QR", field.barcode_type)
#         self.assertEqual("ABC123", field.barcode_value)
#         self.assertEqual("0xF8BD69", field.background_color)
#         self.assertEqual("0xB5413B", field.foreground_color)
#         self.assertEqual("3", field.error_correction_level)
#         self.assertEqual("250", field.scaling_factor)
#         self.assertEqual("1000", field.symbol_height)
#         self.assertEqual("0", field.symbol_rotation)
#
#         field = (FieldDisplayBarcode)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_display_barcode, " DISPLAYBARCODE  501234567890 EAN13 \\t \\p CASE \\x", string.empty, field)
#         self.assertEqual("EAN13", field.barcode_type)
#         self.assertEqual("501234567890", field.barcode_value)
#         self.assertTrue(field.display_text)
#         self.assertEqual("CASE", field.pos_code_style)
#         self.assertTrue(field.fix_check_digit)
#
#         field = (FieldDisplayBarcode)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_display_barcode, " DISPLAYBARCODE  12345ABCDE CODE39 \\d", string.empty, field)
#         self.assertEqual("CODE39", field.barcode_type)
#         self.assertEqual("12345ABCDE", field.barcode_value)
#         self.assertTrue(field.add_start_stop_char)
#
#         field = (FieldDisplayBarcode)doc.range.fields[3]
#
#         TestUtil.verify_field(FieldType.field_display_barcode, " DISPLAYBARCODE  09312345678907 ITF14 \\c STD", string.empty, field)
#         self.assertEqual("ITF14", field.barcode_type)
#         self.assertEqual("09312345678907", field.barcode_value)
#         self.assertEqual("STD", field.case_code_style)
#
#
#     def test_field_merge_barcode___qr(self) :
#
#         #ExStart
#         #ExFor:FieldDisplayBarcode
#         #ExFor:FieldMergeBarcode
#         #ExFor:FieldMergeBarcode.background_color
#         #ExFor:FieldMergeBarcode.barcode_type
#         #ExFor:FieldMergeBarcode.barcode_value
#         #ExFor:FieldMergeBarcode.error_correction_level
#         #ExFor:FieldMergeBarcode.foreground_color
#         #ExFor:FieldMergeBarcode.scaling_factor
#         #ExFor:FieldMergeBarcode.symbol_height
#         #ExFor:FieldMergeBarcode.symbol_rotation
#         #ExSummary:Shows how to perform a mail merge on QR barcodes.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Insert a MERGEBARCODE field, which will accept values from a data source during a mail merge.
#         # This field will convert all values in a merge data source's "MyQRCode" column into QR codes.
#         FieldMergeBarcode field = (FieldMergeBarcode)builder.insert_field(FieldType.field_merge_barcode, true)
#         field.barcode_type = "QR"
#         field.barcode_value = "MyQRCode"
#
#         # Apply custom colors and scaling.
#         field.background_color = "0xF8BD69"
#         field.foreground_color = "0xB5413B"
#         field.error_correction_level = "3"
#         field.scaling_factor = "250"
#         field.symbol_height = "1000"
#         field.symbol_rotation = "0"
#
#         self.assertEqual(FieldType.field_merge_barcode, field.type)
#         self.assertEqual(" MERGEBARCODE  MyQRCode QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0",
#             field.get_field_code())
#         builder.writeln()
#
#         # Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue.
#         # The mail merge will create a new page for each row. Each page will contain a DISPLAYBARCODE field,
#         # which will display a QR code with the value from the merged row.
#         DataTable table = new DataTable("Barcodes")
#         table.columns.add("MyQRCode")
#         table.rows.add(new[]  "ABC123" )
#         table.rows.add(new[]  "DEF456" )
#
#         doc.mail_merge.execute(table)
#
#         self.assertEqual(FieldType.field_display_barcode, doc.range.fields[0].type)
#         self.assertEqual("DISPLAYBARCODE \"ABC123\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B",
#             doc.range.fields[0].get_field_code())
#         self.assertEqual(FieldType.field_display_barcode, doc.range.fields[1].type)
#         self.assertEqual("DISPLAYBARCODE \"DEF456\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B",
#             doc.range.fields[1].get_field_code())
#
#         doc.save(aeb.artifacts_dir + "Field.mergebarcode.qr.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.mergebarcode.qr.docx")
#
#         self.assertEqual(0, doc.range.fields.count(f => f.type == FieldType.field_merge_barcode))
#
#         FieldDisplayBarcode barcode = (FieldDisplayBarcode)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_display_barcode,
#             "DISPLAYBARCODE \"ABC123\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B", string.empty, barcode)
#         self.assertEqual("ABC123", barcode.barcode_value)
#         self.assertEqual("QR", barcode.barcode_type)
#
#         barcode = (FieldDisplayBarcode)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_display_barcode,
#             "DISPLAYBARCODE \"DEF456\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B", string.empty, barcode)
#         self.assertEqual("DEF456", barcode.barcode_value)
#         self.assertEqual("QR", barcode.barcode_type)
#
#
#     def test_field_merge_barcode___ean_13(self) :
#
#         #ExStart
#         #ExFor:FieldMergeBarcode
#         #ExFor:FieldMergeBarcode.barcode_type
#         #ExFor:FieldMergeBarcode.barcode_value
#         #ExFor:FieldMergeBarcode.display_text
#         #ExFor:FieldMergeBarcode.fix_check_digit
#         #ExFor:FieldMergeBarcode.pos_code_style
#         #ExSummary:Shows how to perform a mail merge on EAN13 barcodes.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Insert a MERGEBARCODE field, which will accept values from a data source during a mail merge.
#         # This field will convert all values in a merge data source's "MyEAN13Barcode" column into EAN13 barcodes.
#         FieldMergeBarcode field = (FieldMergeBarcode)builder.insert_field(FieldType.field_merge_barcode, true)
#         field.barcode_type = "EAN13"
#         field.barcode_value = "MyEAN13Barcode"
#
#         # Display the numeric value of the barcode underneath the bars.
#         field.display_text = true
#         field.pos_code_style = "CASE"
#         field.fix_check_digit = true
#
#         self.assertEqual(FieldType.field_merge_barcode, field.type)
#         self.assertEqual(" MERGEBARCODE  MyEAN13Barcode EAN13 \\t \\p CASE \\x", field.get_field_code())
#         builder.writeln()
#
#         # Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue.
#         # The mail merge will create a new page for each row. Each page will contain a DISPLAYBARCODE field,
#         # which will display an EAN13 barcode with the value from the merged row.
#         DataTable table = new DataTable("Barcodes")
#         table.columns.add("MyEAN13Barcode")
#         table.rows.add(new[]  "501234567890" )
#         table.rows.add(new[]  "123456789012" )
#
#         doc.mail_merge.execute(table)
#
#         self.assertEqual(FieldType.field_display_barcode, doc.range.fields[0].type)
#         self.assertEqual("DISPLAYBARCODE \"501234567890\" EAN13 \\t \\p CASE \\x",
#             doc.range.fields[0].get_field_code())
#         self.assertEqual(FieldType.field_display_barcode, doc.range.fields[1].type)
#         self.assertEqual("DISPLAYBARCODE \"123456789012\" EAN13 \\t \\p CASE \\x",
#             doc.range.fields[1].get_field_code())
#
#         doc.save(aeb.artifacts_dir + "Field.mergebarcode.ean_13.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.mergebarcode.ean_13.docx")
#
#         self.assertEqual(0, doc.range.fields.count(f => f.type == FieldType.field_merge_barcode))
#
#         FieldDisplayBarcode barcode = (FieldDisplayBarcode)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_display_barcode, "DISPLAYBARCODE \"501234567890\" EAN13 \\t \\p CASE \\x", string.empty, barcode)
#         self.assertEqual("501234567890", barcode.barcode_value)
#         self.assertEqual("EAN13", barcode.barcode_type)
#
#         barcode = (FieldDisplayBarcode)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_display_barcode, "DISPLAYBARCODE \"123456789012\" EAN13 \\t \\p CASE \\x", string.empty, barcode)
#         self.assertEqual("123456789012", barcode.barcode_value)
#         self.assertEqual("EAN13", barcode.barcode_type)
#
#
#     def test_field_merge_barcode___code_39(self) :
#
#         #ExStart
#         #ExFor:FieldMergeBarcode
#         #ExFor:FieldMergeBarcode.add_start_stop_char
#         #ExFor:FieldMergeBarcode.barcode_type
#         #ExSummary:Shows how to perform a mail merge on CODE39 barcodes.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Insert a MERGEBARCODE field, which will accept values from a data source during a mail merge.
#         # This field will convert all values in a merge data source's "MyCODE39Barcode" column into CODE39 barcodes.
#         FieldMergeBarcode field = (FieldMergeBarcode)builder.insert_field(FieldType.field_merge_barcode, true)
#         field.barcode_type = "CODE39"
#         field.barcode_value = "MyCODE39Barcode"
#
#         # Edit its appearance to display start/stop characters.
#         field.add_start_stop_char = true
#
#         self.assertEqual(FieldType.field_merge_barcode, field.type)
#         self.assertEqual(" MERGEBARCODE  MyCODE39Barcode CODE39 \\d", field.get_field_code())
#         builder.writeln()
#
#         # Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue.
#         # The mail merge will create a new page for each row. Each page will contain a DISPLAYBARCODE field,
#         # which will display a CODE39 barcode with the value from the merged row.
#         DataTable table = new DataTable("Barcodes")
#         table.columns.add("MyCODE39Barcode")
#         table.rows.add(new[]  "12345ABCDE" )
#         table.rows.add(new[]  "67890FGHIJ" )
#
#         doc.mail_merge.execute(table)
#
#         self.assertEqual(FieldType.field_display_barcode, doc.range.fields[0].type)
#         self.assertEqual("DISPLAYBARCODE \"12345ABCDE\" CODE39 \\d",
#             doc.range.fields[0].get_field_code())
#         self.assertEqual(FieldType.field_display_barcode, doc.range.fields[1].type)
#         self.assertEqual("DISPLAYBARCODE \"67890FGHIJ\" CODE39 \\d",
#             doc.range.fields[1].get_field_code())
#
#         doc.save(aeb.artifacts_dir + "Field.mergebarcode.code_39.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.mergebarcode.code_39.docx")
#
#         self.assertEqual(0, doc.range.fields.count(f => f.type == FieldType.field_merge_barcode))
#
#         FieldDisplayBarcode barcode = (FieldDisplayBarcode)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_display_barcode, "DISPLAYBARCODE \"12345ABCDE\" CODE39 \\d", string.empty, barcode)
#         self.assertEqual("12345ABCDE", barcode.barcode_value)
#         self.assertEqual("CODE39", barcode.barcode_type)
#
#         barcode = (FieldDisplayBarcode)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_display_barcode, "DISPLAYBARCODE \"67890FGHIJ\" CODE39 \\d", string.empty, barcode)
#         self.assertEqual("67890FGHIJ", barcode.barcode_value)
#         self.assertEqual("CODE39", barcode.barcode_type)
#
#
#     def test_field_merge_barcode___itf_14(self) :
#
#         #ExStart
#         #ExFor:FieldMergeBarcode
#         #ExFor:FieldMergeBarcode.barcode_type
#         #ExFor:FieldMergeBarcode.case_code_style
#         #ExSummary:Shows how to perform a mail merge on ITF14 barcodes.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Insert a MERGEBARCODE field, which will accept values from a data source during a mail merge.
#         # This field will convert all values in a merge data source's "MyITF14Barcode" column into ITF14 barcodes.
#         FieldMergeBarcode field = (FieldMergeBarcode)builder.insert_field(FieldType.field_merge_barcode, true)
#         field.barcode_type = "ITF14"
#         field.barcode_value = "MyITF14Barcode"
#         field.case_code_style = "STD"
#
#         self.assertEqual(FieldType.field_merge_barcode, field.type)
#         self.assertEqual(" MERGEBARCODE  MyITF14Barcode ITF14 \\c STD", field.get_field_code())
#
#         # Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue.
#         # The mail merge will create a new page for each row. Each page will contain a DISPLAYBARCODE field,
#         # which will display an ITF14 barcode with the value from the merged row.
#         DataTable table = new DataTable("Barcodes")
#         table.columns.add("MyITF14Barcode")
#         table.rows.add(new[]  "09312345678907" )
#         table.rows.add(new[]  "1234567891234" )
#
#         doc.mail_merge.execute(table)
#
#         self.assertEqual(FieldType.field_display_barcode, doc.range.fields[0].type)
#         self.assertEqual("DISPLAYBARCODE \"09312345678907\" ITF14 \\c STD",
#             doc.range.fields[0].get_field_code())
#         self.assertEqual(FieldType.field_display_barcode, doc.range.fields[1].type)
#         self.assertEqual("DISPLAYBARCODE \"1234567891234\" ITF14 \\c STD",
#             doc.range.fields[1].get_field_code())
#
#         doc.save(aeb.artifacts_dir + "Field.mergebarcode.itf_14.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.mergebarcode.itf_14.docx")
#
#         self.assertEqual(0, doc.range.fields.count(f => f.type == FieldType.field_merge_barcode))
#
#         FieldDisplayBarcode barcode = (FieldDisplayBarcode)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_display_barcode, "DISPLAYBARCODE \"09312345678907\" ITF14 \\c STD", string.empty, barcode)
#         self.assertEqual("09312345678907", barcode.barcode_value)
#         self.assertEqual("ITF14", barcode.barcode_type)
#
#         barcode = (FieldDisplayBarcode)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_display_barcode, "DISPLAYBARCODE \"1234567891234\" ITF14 \\c STD", string.empty, barcode)
#         self.assertEqual("1234567891234", barcode.barcode_value)
#         self.assertEqual("ITF14", barcode.barcode_type)
#
#
#     #ExStart
#     #ExFor:FieldLink
#     #ExFor:FieldLink.auto_update
#     #ExFor:FieldLink.format_update_type
#     #ExFor:FieldLink.insert_as_bitmap
#     #ExFor:FieldLink.insert_as_html
#     #ExFor:FieldLink.insert_as_picture
#     #ExFor:FieldLink.insert_as_rtf
#     #ExFor:FieldLink.insert_as_text
#     #ExFor:FieldLink.insert_as_unicode
#     #ExFor:FieldLink.is_linked
#     #ExFor:FieldLink.prog_id
#     #ExFor:FieldLink.source_full_name
#     #ExFor:FieldLink.source_item
#     #ExFor:FieldDde
#     #ExFor:FieldDde.auto_update
#     #ExFor:FieldDde.insert_as_bitmap
#     #ExFor:FieldDde.insert_as_html
#     #ExFor:FieldDde.insert_as_picture
#     #ExFor:FieldDde.insert_as_rtf
#     #ExFor:FieldDde.insert_as_text
#     #ExFor:FieldDde.insert_as_unicode
#     #ExFor:FieldDde.is_linked
#     #ExFor:FieldDde.prog_id
#     #ExFor:FieldDde.source_full_name
#     #ExFor:FieldDde.source_item
#     #ExFor:FieldDdeAuto
#     #ExFor:FieldDdeAuto.insert_as_bitmap
#     #ExFor:FieldDdeAuto.insert_as_html
#     #ExFor:FieldDdeAuto.insert_as_picture
#     #ExFor:FieldDdeAuto.insert_as_rtf
#     #ExFor:FieldDdeAuto.insert_as_text
#     #ExFor:FieldDdeAuto.insert_as_unicode
#     #ExFor:FieldDdeAuto.is_linked
#     #ExFor:FieldDdeAuto.prog_id
#     #ExFor:FieldDdeAuto.source_full_name
#     #ExFor:FieldDdeAuto.source_item
#     #ExSummary:Shows how to use various field types to link to other documents in the local file system, and display their contents.
#     [TestCase(InsertLinkedObjectAs.text)] #ExSkip
#     [TestCase(InsertLinkedObjectAs.unicode)] #ExSkip
#     [TestCase(InsertLinkedObjectAs.html)] #ExSkip
#     [TestCase(InsertLinkedObjectAs.rtf)] #ExSkip
#     [Ignore("WORDSNET-16226")] #ExSkip
#     public void FieldLinkedObjectsAsText(InsertLinkedObjectAs insertLinkedObjectAs)
#
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Below are three types of fields we can use to display contents from a linked document in the form of text.
#         # 1 -  A LINK field:
#         builder.writeln("FieldLink:\n")
#         InsertFieldLink(builder, insertLinkedObjectAs, "Word.document.8", aeb.my_dir + "Document.docx", null, true)
#
#         # 2 -  A DDE field:
#         builder.writeln("FieldDde:\n")
#         InsertFieldDde(builder, insertLinkedObjectAs, "Excel.sheet", aeb.my_dir + "Spreadsheet.xlsx",
#             "Sheet1!R1C1", true, true)
#
#         # 3 -  A DDEAUTO field:
#         builder.writeln("FieldDdeAuto:\n")
#         InsertFieldDdeAuto(builder, insertLinkedObjectAs, "Excel.sheet", aeb.my_dir + "Spreadsheet.xlsx",
#             "Sheet1!R1C1", true)
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.link.dde.ddeauto.docx")
#
#
#     [TestCase(InsertLinkedObjectAs.picture)] #ExSkip
#     [TestCase(InsertLinkedObjectAs.bitmap)] #ExSkip
#     [Ignore("WORDSNET-16226")] #ExSkip
#     public void FieldLinkedObjectsAsImage(InsertLinkedObjectAs insertLinkedObjectAs)
#
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Below are three types of fields we can use to display contents from a linked document in the form of an image.
#         # 1 -  A LINK field:
#         builder.writeln("FieldLink:\n")
#         InsertFieldLink(builder, insertLinkedObjectAs, "Excel.sheet", aeb.my_dir + "MySpreadsheet.xlsx",
#             "Sheet1!R2C2", true)
#
#         # 2 -  A DDE field:
#         builder.writeln("FieldDde:\n")
#         InsertFieldDde(builder, insertLinkedObjectAs, "Excel.sheet", aeb.my_dir + "Spreadsheet.xlsx",
#             "Sheet1!R1C1", true, true)
#
#         # 3 -  A DDEAUTO field:
#         builder.writeln("FieldDdeAuto:\n")
#         InsertFieldDdeAuto(builder, insertLinkedObjectAs, "Excel.sheet", aeb.my_dir + "Spreadsheet.xlsx",
#             "Sheet1!R1C1", true)
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.link.dde.ddeauto.as_image.docx")
#
#
#     # <summary>
#     # Use a document builder to insert a LINK field and set its properties according to parameters.
#     # </summary>
#     private static void InsertFieldLink(DocumentBuilder builder, InsertLinkedObjectAs insertLinkedObjectAs,
#         string progId, string sourceFullName, string sourceItem, bool shouldAutoUpdate)
#
#         FieldLink field = (FieldLink)builder.insert_field(FieldType.field_link, true)
#
#         switch (insertLinkedObjectAs)
#
#             case InsertLinkedObjectAs.text:
#                 field.insert_as_text = true
#                 break
#             case InsertLinkedObjectAs.unicode:
#                 field.insert_as_unicode = true
#                 break
#             case InsertLinkedObjectAs.html:
#                 field.insert_as_html = true
#                 break
#             case InsertLinkedObjectAs.rtf:
#                 field.insert_as_rtf = true
#                 break
#             case InsertLinkedObjectAs.picture:
#                 field.insert_as_picture = true
#                 break
#             case InsertLinkedObjectAs.bitmap:
#                 field.insert_as_bitmap = true
#                 break
#
#
#         field.auto_update = shouldAutoUpdate
#         field.prog_id = progId
#         field.source_full_name = sourceFullName
#         field.source_item = sourceItem
#
#         builder.writeln("\n")
#
#
#     # <summary>
#     # Use a document builder to insert a DDE field, and set its properties according to parameters.
#     # </summary>
#     private static void InsertFieldDde(DocumentBuilder builder, InsertLinkedObjectAs insertLinkedObjectAs, string progId,
#         string sourceFullName, string sourceItem, bool isLinked, bool shouldAutoUpdate)
#
#         FieldDde field = (FieldDde)builder.insert_field(FieldType.field_dde, true)
#
#         switch (insertLinkedObjectAs)
#
#             case InsertLinkedObjectAs.text:
#                 field.insert_as_text = true
#                 break
#             case InsertLinkedObjectAs.unicode:
#                 field.insert_as_unicode = true
#                 break
#             case InsertLinkedObjectAs.html:
#                 field.insert_as_html = true
#                 break
#             case InsertLinkedObjectAs.rtf:
#                 field.insert_as_rtf = true
#                 break
#             case InsertLinkedObjectAs.picture:
#                 field.insert_as_picture = true
#                 break
#             case InsertLinkedObjectAs.bitmap:
#                 field.insert_as_bitmap = true
#                 break
#
#
#         field.auto_update = shouldAutoUpdate
#         field.prog_id = progId
#         field.source_full_name = sourceFullName
#         field.source_item = sourceItem
#         field.is_linked = isLinked
#
#         builder.writeln("\n")
#
#
#     # <summary>
#     # Use a document builder to insert a DDEAUTO, field and set its properties according to parameters.
#     # </summary>
#     private static void InsertFieldDdeAuto(DocumentBuilder builder, InsertLinkedObjectAs insertLinkedObjectAs,
#         string progId, string sourceFullName, string sourceItem, bool isLinked)
#
#         FieldDdeAuto field = (FieldDdeAuto)builder.insert_field(FieldType.field_dde_auto, true)
#
#         switch (insertLinkedObjectAs)
#
#             case InsertLinkedObjectAs.text:
#                 field.insert_as_text = true
#                 break
#             case InsertLinkedObjectAs.unicode:
#                 field.insert_as_unicode = true
#                 break
#             case InsertLinkedObjectAs.html:
#                 field.insert_as_html = true
#                 break
#             case InsertLinkedObjectAs.rtf:
#                 field.insert_as_rtf = true
#                 break
#             case InsertLinkedObjectAs.picture:
#                 field.insert_as_picture = true
#                 break
#             case InsertLinkedObjectAs.bitmap:
#                 field.insert_as_bitmap = true
#                 break
#
#
#         field.prog_id = progId
#         field.source_full_name = sourceFullName
#         field.source_item = sourceItem
#         field.is_linked = isLinked
#
#
#     public enum InsertLinkedObjectAs
#
#         # LinkedObjectAsText
#         Text,
#         Unicode,
#         Html,
#         Rtf,
#         # LinkedObjectAsImage
#         Picture,
#         Bitmap
#
#     #ExEnd
#
#     def test_field_user_address(self) :
#
#         #ExStart
#         #ExFor:FieldUserAddress
#         #ExFor:FieldUserAddress.user_address
#         #ExSummary:Shows how to use the USERADDRESS field.
#         doc = aw.Document()
#
#         # Create a UserInformation object and set it as the source of user information for any fields that we create.
#         UserInformation userInformation = new UserInformation()
#         userInformation.address = "123 Main Street"
#         doc.field_options.current_user = userInformation
#
#         # Create a USERADDRESS field to display the current user's address,
#         # taken from the UserInformation object we created above.
#         builder = aw.DocumentBuilder(doc)
#         FieldUserAddress fieldUserAddress = (FieldUserAddress)builder.insert_field(FieldType.field_user_address, true)
#         self.assertEqual(userInformation.address, fieldUserAddress.result) #ExSkip
#
#         self.assertEqual(" USERADDRESS ", fieldUserAddress.get_field_code())
#         self.assertEqual("123 Main Street", fieldUserAddress.result)
#
#         # We can set this property to get our field to override the value currently stored in the UserInformation object.
#         fieldUserAddress.user_address = "456 North Road"
#         fieldUserAddress.update()
#
#         self.assertEqual(" USERADDRESS  \"456 North Road\"", fieldUserAddress.get_field_code())
#         self.assertEqual("456 North Road", fieldUserAddress.result)
#
#         # This does not affect the value in the UserInformation object.
#         self.assertEqual("123 Main Street", doc.field_options.current_user.address)
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.useraddress.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.useraddress.docx")
#
#         fieldUserAddress = (FieldUserAddress)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_user_address, " USERADDRESS  \"456 North Road\"", "456 North Road", fieldUserAddress)
#         self.assertEqual("456 North Road", fieldUserAddress.user_address)
#
#
#     def test_field_user_initials(self) :
#
#         #ExStart
#         #ExFor:FieldUserInitials
#         #ExFor:FieldUserInitials.user_initials
#         #ExSummary:Shows how to use the USERINITIALS field.
#         doc = aw.Document()
#
#         # Create a UserInformation object and set it as the source of user information for any fields that we create.
#         UserInformation userInformation = new UserInformation()
#         userInformation.initials = "J. D."
#         doc.field_options.current_user = userInformation
#
#         # Create a USERINITIALS field to display the current user's initials,
#         # taken from the UserInformation object we created above.
#         builder = aw.DocumentBuilder(doc)
#         FieldUserInitials fieldUserInitials = (FieldUserInitials)builder.insert_field(FieldType.field_user_initials, true)
#         self.assertEqual(userInformation.initials, fieldUserInitials.result)
#
#         self.assertEqual(" USERINITIALS ", fieldUserInitials.get_field_code())
#         self.assertEqual("J. D.", fieldUserInitials.result)
#
#         # We can set this property to get our field to override the value currently stored in the UserInformation object.
#         fieldUserInitials.user_initials = "J. C."
#         fieldUserInitials.update()
#
#         self.assertEqual(" USERINITIALS  \"J. C.\"", fieldUserInitials.get_field_code())
#         self.assertEqual("J. C.", fieldUserInitials.result)
#
#         # This does not affect the value in the UserInformation object.
#         self.assertEqual("J. D.", doc.field_options.current_user.initials)
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.userinitials.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.userinitials.docx")
#
#         fieldUserInitials = (FieldUserInitials)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_user_initials, " USERINITIALS  \"J. C.\"", "J. C.", fieldUserInitials)
#         self.assertEqual("J. C.", fieldUserInitials.user_initials)
#
#
#     def test_field_user_name(self) :
#
#         #ExStart
#         #ExFor:FieldUserName
#         #ExFor:FieldUserName.user_name
#         #ExSummary:Shows how to use the USERNAME field.
#         doc = aw.Document()
#
#         # Create a UserInformation object and set it as the source of user information for any fields that we create.
#         UserInformation userInformation = new UserInformation()
#         userInformation.name = "John Doe"
#         doc.field_options.current_user = userInformation
#
#         builder = aw.DocumentBuilder(doc)
#
#         # Create a USERNAME field to display the current user's name,
#         # taken from the UserInformation object we created above.
#         FieldUserName fieldUserName = (FieldUserName)builder.insert_field(FieldType.field_user_name, true)
#         self.assertEqual(userInformation.name, fieldUserName.result)
#
#         self.assertEqual(" USERNAME ", fieldUserName.get_field_code())
#         self.assertEqual("John Doe", fieldUserName.result)
#
#         # We can set this property to get our field to override the value currently stored in the UserInformation object.
#         fieldUserName.user_name = "Jane Doe"
#         fieldUserName.update()
#
#         self.assertEqual(" USERNAME  \"Jane Doe\"", fieldUserName.get_field_code())
#         self.assertEqual("Jane Doe", fieldUserName.result)
#
#         # This does not affect the value in the UserInformation object.
#         self.assertEqual("John Doe", doc.field_options.current_user.name)
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.username.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.username.docx")
#
#         fieldUserName = (FieldUserName)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_user_name, " USERNAME  \"Jane Doe\"", "Jane Doe", fieldUserName)
#         self.assertEqual("Jane Doe", fieldUserName.user_name)
#
#
#     [Test]
#     [Ignore("WORDSNET-17657")]
#     public void FieldStyleRefParagraphNumbers()
#
#         #ExStart
#         #ExFor:FieldStyleRef
#         #ExFor:FieldStyleRef.insert_paragraph_number
#         #ExFor:FieldStyleRef.insert_paragraph_number_in_full_context
#         #ExFor:FieldStyleRef.insert_paragraph_number_in_relative_context
#         #ExFor:FieldStyleRef.insert_relative_position
#         #ExFor:FieldStyleRef.search_from_bottom
#         #ExFor:FieldStyleRef.style_name
#         #ExFor:FieldStyleRef.suppress_non_delimiters
#         #ExSummary:Shows how to use STYLEREF fields.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Create a list based using a Microsoft Word list template.
#         Aspose.words.lists.list list = doc.lists.add(Aspose.words.lists.list_template.number_default)
#
#         # This generated list will display "1.a )".
#         # Space before the bracket is a non-delimiter character, which we can suppress.
#         list.list_levels[0].number_format = "\x0000."
#         list.list_levels[1].number_format = "\x0001 )"
#
#         # Add text and apply paragraph styles that STYLEREF fields will reference.
#         builder.list_format.list = list
#         builder.list_format.list_indent()
#         builder.paragraph_format.style = doc.styles["List Paragraph"]
#         builder.writeln("Item 1")
#         builder.paragraph_format.style = doc.styles["Quote"]
#         builder.writeln("Item 2")
#         builder.paragraph_format.style = doc.styles["List Paragraph"]
#         builder.writeln("Item 3")
#         builder.list_format.remove_numbers()
#         builder.paragraph_format.style = doc.styles["Normal"]
#
#         # Place a STYLEREF field in the header and display the first "List Paragraph"-styled text in the document.
#         builder.move_to_header_footer(HeaderFooterType.header_primary)
#         FieldStyleRef field = (FieldStyleRef)builder.insert_field(FieldType.field_style_ref, true)
#         field.style_name = "List Paragraph"
#
#         # Place a STYLEREF field in the footer, and have it display the last text.
#         builder.move_to_header_footer(HeaderFooterType.footer_primary)
#         field = (FieldStyleRef)builder.insert_field(FieldType.field_style_ref, true)
#         field.style_name = "List Paragraph"
#         field.search_from_bottom = true
#
#         builder.move_to_document_end()
#
#         # We can also use STYLEREF fields to reference the list numbers of lists.
#         builder.write("\nParagraph number: ")
#         field = (FieldStyleRef)builder.insert_field(FieldType.field_style_ref, true)
#         field.style_name = "Quote"
#         field.insert_paragraph_number = true
#
#         builder.write("\nParagraph number, relative context: ")
#         field = (FieldStyleRef)builder.insert_field(FieldType.field_style_ref, true)
#         field.style_name = "Quote"
#         field.insert_paragraph_number_in_relative_context = true
#
#         builder.write("\nParagraph number, full context: ")
#         field = (FieldStyleRef)builder.insert_field(FieldType.field_style_ref, true)
#         field.style_name = "Quote"
#         field.insert_paragraph_number_in_full_context = true
#
#         builder.write("\nParagraph number, full context, non-delimiter chars suppressed: ")
#         field = (FieldStyleRef)builder.insert_field(FieldType.field_style_ref, true)
#         field.style_name = "Quote"
#         field.insert_paragraph_number_in_full_context = true
#         field.suppress_non_delimiters = true
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.styleref.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.styleref.docx")
#
#         field = (FieldStyleRef)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_style_ref, " STYLEREF  \"List Paragraph\"", "Item 1", field)
#         self.assertEqual("List Paragraph", field.style_name)
#
#         field = (FieldStyleRef)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_style_ref, " STYLEREF  \"List Paragraph\" \\l", "Item 3", field)
#         self.assertEqual("List Paragraph", field.style_name)
#         self.assertTrue(field.search_from_bottom)
#
#         field = (FieldStyleRef)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_style_ref, " STYLEREF  Quote \\n", "b )", field)
#         self.assertEqual("Quote", field.style_name)
#         self.assertTrue(field.insert_paragraph_number)
#
#         field = (FieldStyleRef)doc.range.fields[3]
#
#         TestUtil.verify_field(FieldType.field_style_ref, " STYLEREF  Quote \\r", "b )", field)
#         self.assertEqual("Quote", field.style_name)
#         self.assertTrue(field.insert_paragraph_number_in_relative_context)
#
#         field = (FieldStyleRef)doc.range.fields[4]
#
#         TestUtil.verify_field(FieldType.field_style_ref, " STYLEREF  Quote \\w", "1.b )", field)
#         self.assertEqual("Quote", field.style_name)
#         self.assertTrue(field.insert_paragraph_number_in_full_context)
#
#         field = (FieldStyleRef)doc.range.fields[5]
#
#         TestUtil.verify_field(FieldType.field_style_ref, " STYLEREF  Quote \\w \\t", "1.b)", field)
#         self.assertEqual("Quote", field.style_name)
#         self.assertTrue(field.insert_paragraph_number_in_full_context)
#         self.assertTrue(field.suppress_non_delimiters)
#
#
# #if NET462 || NETCOREAPP2_1 || JAVA
#     def test_field_date(self) :
#
#         #ExStart
#         #ExFor:FieldDate
#         #ExFor:FieldDate.use_lunar_calendar
#         #ExFor:FieldDate.use_saka_era_calendar
#         #ExFor:FieldDate.use_um_al_qura_calendar
#         #ExFor:FieldDate.use_last_format
#         #ExSummary:Shows how to use DATE fields to display dates according to different kinds of calendars.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # If we want the text in the document always to display the correct date, we can use a DATE field.
#         # Below are three types of cultural calendars that a DATE field can use to display a date.
#         # 1 -  Islamic Lunar Calendar:
#         FieldDate field = (FieldDate)builder.insert_field(FieldType.field_date, true)
#         field.use_lunar_calendar = true
#         self.assertEqual(" DATE  \\h", field.get_field_code())
#         builder.writeln()
#
#         # 2 -  Umm al-Qura calendar:
#         field = (FieldDate)builder.insert_field(FieldType.field_date, true)
#         field.use_um_al_qura_calendar = true
#         self.assertEqual(" DATE  \\u", field.get_field_code())
#         builder.writeln()
#
#         # 3 -  Indian National Calendar:
#         field = (FieldDate)builder.insert_field(FieldType.field_date, true)
#         field.use_saka_era_calendar = true
#         self.assertEqual(" DATE  \\s", field.get_field_code())
#         builder.writeln()
#
#         # Insert a DATE field and set its calendar type to the one last used by the host application.
#         # In Microsoft Word, the type will be the most recently used in the Insert -> Text -> Date and Time dialog box.
#         field = (FieldDate)builder.insert_field(FieldType.field_date, true)
#         field.use_last_format = true
#         self.assertEqual(" DATE  \\l", field.get_field_code())
#         builder.writeln()
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.date.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.date.docx")
#
#         field = (FieldDate)doc.range.fields[0]
#
#         self.assertEqual(FieldType.field_date, field.type)
#         self.assertTrue(field.use_lunar_calendar)
#         self.assertEqual(" DATE  \\h", field.get_field_code())
#         self.assertTrue(Regex.match(doc.range.fields[0].result, @"\d1,2[/]\d1,2[/]\d4").success)
#
#         field = (FieldDate)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_date, " DATE  \\u", DateTime.now.to_short_date_string(), field)
#         self.assertTrue(field.use_um_al_qura_calendar)
#
#         field = (FieldDate)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_date, " DATE  \\s", DateTime.now.to_short_date_string(), field)
#         self.assertTrue(field.use_saka_era_calendar)
#
#         field = (FieldDate)doc.range.fields[3]
#
#         TestUtil.verify_field(FieldType.field_date, " DATE  \\l", DateTime.now.to_short_date_string(), field)
#         self.assertTrue(field.use_last_format)
#
# #endif
#
#     [Test]
#     [Ignore("WORDSNET-17669")]
#     public void FieldCreateDate()
#
#         #ExStart
#         #ExFor:FieldCreateDate
#         #ExFor:FieldCreateDate.use_lunar_calendar
#         #ExFor:FieldCreateDate.use_saka_era_calendar
#         #ExFor:FieldCreateDate.use_um_al_qura_calendar
#         #ExSummary:Shows how to use the CREATEDATE field to display the creation date/time of the document.
#         Document doc = new Document(aeb.my_dir + "Document.docx")
#         builder = aw.DocumentBuilder(doc)
#         builder.move_to_document_end()
#         builder.writeln(" Date this document was created:")
#
#         # We can use the CREATEDATE field to display the date and time of the creation of the document.
#         # Below are three different calendar types according to which the CREATEDATE field can display the date/time.
#         # 1 -  Islamic Lunar Calendar:
#         builder.write("According to the Lunar Calendar - ")
#         FieldCreateDate field = (FieldCreateDate)builder.insert_field(FieldType.field_create_date, true)
#         field.use_lunar_calendar = true
#
#         self.assertEqual(" CREATEDATE  \\h", field.get_field_code())
#
#         # 2 -  Umm al-Qura calendar:
#         builder.write("\nAccording to the Umm al-Qura Calendar - ")
#         field = (FieldCreateDate)builder.insert_field(FieldType.field_create_date, true)
#         field.use_um_al_qura_calendar = true
#
#         self.assertEqual(" CREATEDATE  \\u", field.get_field_code())
#
#         # 3 -  Indian National Calendar:
#         builder.write("\nAccording to the Indian National Calendar - ")
#         field = (FieldCreateDate)builder.insert_field(FieldType.field_create_date, true)
#         field.use_saka_era_calendar = true
#
#         self.assertEqual(" CREATEDATE  \\s", field.get_field_code())
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.createdate.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.createdate.docx")
#
#         self.assertEqual(new DateTime(2017, 12, 5, 9, 56, 0), doc.built_in_document_properties.created_time)
#
#         DateTime expectedDate = doc.built_in_document_properties.created_time.add_hours(TimeZoneInfo.local.get_utc_offset(DateTime.utc_now).hours)
#         field = (FieldCreateDate)doc.range.fields[0]
#         Calendar umAlQuraCalendar = new UmAlQuraCalendar()
#
#         TestUtil.verify_field(FieldType.field_create_date, " CREATEDATE  \\h",
#             $"umAlQuraCalendar.get_month(expectedDate)/umAlQuraCalendar.get_day_of_month(expectedDate)/umAlQuraCalendar.get_year(expectedDate) " +
#             expectedDate.add_hours(1).to_string("hh:mm:ss tt"), field)
#         self.assertEqual(FieldType.field_create_date, field.type)
#         self.assertTrue(field.use_lunar_calendar)
#
#         field = (FieldCreateDate)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_create_date, " CREATEDATE  \\u",
#             $"umAlQuraCalendar.get_month(expectedDate)/umAlQuraCalendar.get_day_of_month(expectedDate)/umAlQuraCalendar.get_year(expectedDate) " +
#             expectedDate.add_hours(1).to_string("hh:mm:ss tt"), field)
#         self.assertEqual(FieldType.field_create_date, field.type)
#         self.assertTrue(field.use_um_al_qura_calendar)
#
#
#     [Test]
#     [Ignore("WORDSNET-17669")]
#     public void FieldSaveDate()
#
#         #ExStart
#         #ExFor:BuiltInDocumentProperties.last_saved_time
#         #ExFor:FieldSaveDate
#         #ExFor:FieldSaveDate.use_lunar_calendar
#         #ExFor:FieldSaveDate.use_saka_era_calendar
#         #ExFor:FieldSaveDate.use_um_al_qura_calendar
#         #ExSummary:Shows how to use the SAVEDATE field to display the date/time of the document's most recent save operation performed using Microsoft Word.
#         Document doc = new Document(aeb.my_dir + "Document.docx")
#         builder = aw.DocumentBuilder(doc)
#         builder.move_to_document_end()
#         builder.writeln(" Date this document was last saved:")
#
#         # We can use the SAVEDATE field to display the last save operation's date and time on the document.
#         # The save operation that these fields refer to is the manual save in an application like Microsoft Word,
#         # not the document's Save method.
#         # Below are three different calendar types according to which the SAVEDATE field can display the date/time.
#         # 1 -  Islamic Lunar Calendar:
#         builder.write("According to the Lunar Calendar - ")
#         FieldSaveDate field = (FieldSaveDate)builder.insert_field(FieldType.field_save_date, true)
#         field.use_lunar_calendar = true
#
#         self.assertEqual(" SAVEDATE  \\h", field.get_field_code())
#
#         # 2 -  Umm al-Qura calendar:
#         builder.write("\nAccording to the Umm al-Qura calendar - ")
#         field = (FieldSaveDate)builder.insert_field(FieldType.field_save_date, true)
#         field.use_um_al_qura_calendar = true
#
#         self.assertEqual(" SAVEDATE  \\u", field.get_field_code())
#
#         # 3 -  Indian National calendar:
#         builder.write("\nAccording to the Indian National calendar - ")
#         field = (FieldSaveDate)builder.insert_field(FieldType.field_save_date, true)
#         field.use_saka_era_calendar = true
#
#         self.assertEqual(" SAVEDATE  \\s", field.get_field_code())
#
#         # The SAVEDATE fields draw their date/time values from the LastSavedTime built-in property.
#         # The document's Save method will not update this value, but we can still update it manually.
#         doc.built_in_document_properties.last_saved_time = DateTime.now
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.savedate.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.savedate.docx")
#
#         print(doc.built_in_document_properties.last_saved_time)
#
#         field = (FieldSaveDate)doc.range.fields[0]
#
#         self.assertEqual(FieldType.field_save_date, field.type)
#         self.assertTrue(field.use_lunar_calendar)
#         self.assertEqual(" SAVEDATE  \\h", field.get_field_code())
#
#         self.assertTrue(Regex.match(field.result, "\\d1,2[/]\\d1,2[/]\\d4 \\d1,2:\\d1,2:\\d1,2 [A,P]M").success)
#
#         field = (FieldSaveDate)doc.range.fields[1]
#
#         self.assertEqual(FieldType.field_save_date, field.type)
#         self.assertTrue(field.use_um_al_qura_calendar)
#         self.assertEqual(" SAVEDATE  \\u", field.get_field_code())
#         self.assertTrue(Regex.match(field.result, "\\d1,2[/]\\d1,2[/]\\d4 \\d1,2:\\d1,2:\\d1,2 [A,P]M").success)
#
#
#     def test_field_builder(self) :
#
#         #ExStart
#         #ExFor:FieldBuilder
#         #ExFor:FieldBuilder.add_argument(Int32)
#         #ExFor:FieldBuilder.add_argument(FieldArgumentBuilder)
#         #ExFor:FieldBuilder.add_argument(String)
#         #ExFor:FieldBuilder.add_argument(Double)
#         #ExFor:FieldBuilder.add_argument(FieldBuilder)
#         #ExFor:FieldBuilder.add_switch(String)
#         #ExFor:FieldBuilder.add_switch(String, Double)
#         #ExFor:FieldBuilder.add_switch(String, Int32)
#         #ExFor:FieldBuilder.add_switch(String, String)
#         #ExFor:FieldBuilder.build_and_insert(Paragraph)
#         #ExFor:FieldArgumentBuilder
#         #ExFor:FieldArgumentBuilder.add_field(FieldBuilder)
#         #ExFor:FieldArgumentBuilder.add_text(String)
#         #ExFor:FieldArgumentBuilder.add_node(Inline)
#         #ExSummary:Shows how to construct fields using a field builder, and then insert them into the document.
#         doc = aw.Document()
#
#         # Below are three examples of field construction done using a field builder.
#         # 1 -  Single field:
#         # Use a field builder to add a SYMBOL field which displays the ƒ (Florin) symbol.
#         FieldBuilder builder = new FieldBuilder(FieldType.field_symbol)
#         builder.add_argument(402)
#         builder.add_switch("\\f", "Arial")
#         builder.add_switch("\\s", 25)
#         builder.add_switch("\\u")
#         Field field = builder.build_and_insert(doc.first_section.body.first_paragraph)
#
#         self.assertEqual(" SYMBOL 402 \\f Arial \\s 25 \\u ", field.get_field_code())
#
#         # 2 -  Nested field:
#         # Use a field builder to create a formula field used as an inner field by another field builder.
#         FieldBuilder innerFormulaBuilder = new FieldBuilder(FieldType.field_formula)
#         innerFormulaBuilder.add_argument(100)
#         innerFormulaBuilder.add_argument("+")
#         innerFormulaBuilder.add_argument(74)
#
#         # Create another builder for another SYMBOL field, and insert the formula field
#         # that we have created above into the SYMBOL field as its argument.
#         builder = new FieldBuilder(FieldType.field_symbol)
#         builder.add_argument(innerFormulaBuilder)
#         field = builder.build_and_insert(doc.first_section.body.append_paragraph(string.empty))
#
#         # The outer SYMBOL field will use the formula field result, 174, as its argument,
#         # which will make the field display the ® (Registered Sign) symbol since its character number is 174.
#         self.assertEqual(" SYMBOL \u0013 = 100 + 74 \u0014\u0015 ", field.get_field_code())
#
#         # 3 -  Multiple nested fields and arguments:
#         # Now, we will use a builder to create an IF field, which displays one of two custom string values,
#         # depending on the true/false value of its expression. To get a true/false value
#         # that determines which string the IF field displays, the IF field will test two numeric expressions for equality.
#         # We will provide the two expressions in the form of formula fields, which we will nest inside the IF field.
#         FieldBuilder leftExpression = new FieldBuilder(FieldType.field_formula)
#         leftExpression.add_argument(2)
#         leftExpression.add_argument("+")
#         leftExpression.add_argument(3)
#
#         FieldBuilder rightExpression = new FieldBuilder(FieldType.field_formula)
#         rightExpression.add_argument(2.5)
#         rightExpression.add_argument("*")
#         rightExpression.add_argument(5.2)
#
#         # Next, we will build two field arguments, which will serve as the true/false output strings for the IF field.
#         # These arguments will reuse the output values of our numeric expressions.
#         FieldArgumentBuilder trueOutput = new FieldArgumentBuilder()
#         trueOutput.add_text("True, both expressions amount to ")
#         trueOutput.add_field(leftExpression)
#
#         FieldArgumentBuilder falseOutput = new FieldArgumentBuilder()
#         falseOutput.add_node(new Run(doc, "False, "))
#         falseOutput.add_field(leftExpression)
#         falseOutput.add_node(new Run(doc, " does not equal "))
#         falseOutput.add_field(rightExpression)
#
#         # Finally, we will create one more field builder for the IF field and combine all of the expressions.
#         builder = new FieldBuilder(FieldType.field_if)
#         builder.add_argument(leftExpression)
#         builder.add_argument("=")
#         builder.add_argument(rightExpression)
#         builder.add_argument(trueOutput)
#         builder.add_argument(falseOutput)
#         field = builder.build_and_insert(doc.first_section.body.append_paragraph(string.empty))
#
#         self.assertEqual(" IF \u0013 = 2 + 3 \u0014\u0015 = \u0013 = 2.5 * 5.2 \u0014\u0015 " +
#                         "\"True, both expressions amount to \u0013 = 2 + 3 \u0014\u0015\" " +
#                         "\"False, \u0013 = 2 + 3 \u0014\u0015 does not equal \u0013 = 2.5 * 5.2 \u0014\u0015\" ", field.get_field_code())
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.symbol.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.symbol.docx")
#
#         FieldSymbol fieldSymbol = (FieldSymbol)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_symbol, " SYMBOL 402 \\f Arial \\s 25 \\u ", string.empty, fieldSymbol)
#         self.assertEqual("ƒ", fieldSymbol.display_result)
#
#         fieldSymbol = (FieldSymbol)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_symbol, " SYMBOL \u0013 = 100 + 74 \u0014174\u0015 ", string.empty, fieldSymbol)
#         self.assertEqual("®", fieldSymbol.display_result)
#
#         TestUtil.verify_field(FieldType.field_formula, " = 100 + 74 ", "174", doc.range.fields[2])
#
#         TestUtil.verify_field(FieldType.field_if,
#             " IF \u0013 = 2 + 3 \u00145\u0015 = \u0013 = 2.5 * 5.2 \u001413\u0015 " +
#             "\"True, both expressions amount to \u0013 = 2 + 3 \u0014\u0015\" " +
#             "\"False, \u0013 = 2 + 3 \u00145\u0015 does not equal \u0013 = 2.5 * 5.2 \u001413\u0015\" ",
#             "False, 5 does not equal 13", doc.range.fields[3])
#
#         Assert.throws<AssertionException>(() => TestUtil.fields_are_nested(doc.range.fields[2], doc.range.fields[3]))
#
#         TestUtil.verify_field(FieldType.field_formula, " = 2 + 3 ", "5", doc.range.fields[4])
#         TestUtil.fields_are_nested(doc.range.fields[4], doc.range.fields[3])
#
#         TestUtil.verify_field(FieldType.field_formula, " = 2.5 * 5.2 ", "13", doc.range.fields[5])
#         TestUtil.fields_are_nested(doc.range.fields[5], doc.range.fields[3])
#
#         TestUtil.verify_field(FieldType.field_formula, " = 2 + 3 ", string.empty, doc.range.fields[6])
#         TestUtil.fields_are_nested(doc.range.fields[6], doc.range.fields[3])
#
#         TestUtil.verify_field(FieldType.field_formula, " = 2 + 3 ", "5", doc.range.fields[7])
#         TestUtil.fields_are_nested(doc.range.fields[7], doc.range.fields[3])
#
#         TestUtil.verify_field(FieldType.field_formula, " = 2.5 * 5.2 ", "13", doc.range.fields[8])
#         TestUtil.fields_are_nested(doc.range.fields[8], doc.range.fields[3])
#
#
#     def test_field_author(self) :
#
#         #ExStart
#         #ExFor:FieldAuthor
#         #ExFor:FieldAuthor.author_name
#         #ExFor:FieldOptions.default_document_author
#         #ExSummary:Shows how to use an AUTHOR field to display a document creator's name.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # AUTHOR fields source their results from the built-in document property called "Author".
#         # If we create and save a document in Microsoft Word,
#         # it will have our username in that property.
#         # However, if we create a document programmatically using Aspose.words,
#         # the "Author" property, by default, will be an empty string.
#         self.assertEqual(string.empty, doc.built_in_document_properties.author)
#
#         # Set a backup author name for AUTHOR fields to use
#         # if the "Author" property contains an empty string.
#         doc.field_options.default_document_author = "Joe Bloggs"
#
#         builder.write("This document was created by ")
#         FieldAuthor field = (FieldAuthor)builder.insert_field(FieldType.field_author, true)
#         field.update()
#
#         self.assertEqual(" AUTHOR ", field.get_field_code())
#         self.assertEqual("Joe Bloggs", field.result)
#
#         # Updating an AUTHOR field that contains a value
#         # will apply that value to the "Author" built-in property.
#         self.assertEqual("Joe Bloggs", doc.built_in_document_properties.author)
#
#         # Changing this property, then updating the AUTHOR field will apply this value to the field.
#         doc.built_in_document_properties.author = "John Doe"
#         field.update()
#
#         self.assertEqual(" AUTHOR ", field.get_field_code())
#         self.assertEqual("John Doe", field.result)
#
#         # If we update an AUTHOR field after changing its "Name" property,
#         # then the field will display the new name and apply the new name to the built-in property.
#         field.author_name = "Jane Doe"
#         field.update()
#
#         self.assertEqual(" AUTHOR  \"Jane Doe\"", field.get_field_code())
#         self.assertEqual("Jane Doe", field.result)
#
#         # AUTHOR fields do not affect the DefaultDocumentAuthor property.
#         self.assertEqual("Jane Doe", doc.built_in_document_properties.author)
#         self.assertEqual("Joe Bloggs", doc.field_options.default_document_author)
#
#         doc.save(aeb.artifacts_dir + "Field.author.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.author.docx")
#
#         Assert.null(doc.field_options.default_document_author)
#         self.assertEqual("Jane Doe", doc.built_in_document_properties.author)
#
#         field = (FieldAuthor)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_author, " AUTHOR  \"Jane Doe\"", "Jane Doe", field)
#         self.assertEqual("Jane Doe", field.author_name)
#
#
#     def test_field_doc_variable(self) :
#
#         #ExStart
#         #ExFor:FieldDocProperty
#         #ExFor:FieldDocVariable
#         #ExFor:FieldDocVariable.variable_name
#         #ExSummary:Shows how to use DOCPROPERTY fields to display document properties and variables.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Below are two ways of using DOCPROPERTY fields.
#         # 1 -  Display a built-in property:
#         # Set a custom value for the "Category" built-in property, then insert a DOCPROPERTY field that references it.
#         doc.built_in_document_properties.category = "My category"
#
#         FieldDocProperty fieldDocProperty = (FieldDocProperty)builder.insert_field(" DOCPROPERTY Category ")
#         fieldDocProperty.update()
#
#         self.assertEqual(" DOCPROPERTY Category ", fieldDocProperty.get_field_code())
#         self.assertEqual("My category", fieldDocProperty.result)
#
#         builder.insert_paragraph()
#
#         # 2 -  Display a custom document variable:
#         # Define a custom variable, then reference that variable with a DOCPROPERTY field.
#         Assert.that(doc.variables, Is.empty)
#         doc.variables.add("My variable", "My variable's value")
#
#         FieldDocVariable fieldDocVariable = (FieldDocVariable)builder.insert_field(FieldType.field_doc_variable, true)
#         fieldDocVariable.variable_name = "My Variable"
#         fieldDocVariable.update()
#
#         self.assertEqual(" DOCVARIABLE  \"My Variable\"", fieldDocVariable.get_field_code())
#         self.assertEqual("My variable's value", fieldDocVariable.result)
#
#         doc.save(aeb.artifacts_dir + "Field.docproperty.docvariable.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.docproperty.docvariable.docx")
#
#         self.assertEqual("My category", doc.built_in_document_properties.category)
#
#         fieldDocProperty = (FieldDocProperty)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_doc_property, " DOCPROPERTY Category ", "My category", fieldDocProperty)
#
#         fieldDocVariable = (FieldDocVariable)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_doc_variable, " DOCVARIABLE  \"My Variable\"", "My variable's value", fieldDocVariable)
#         self.assertEqual("My Variable", fieldDocVariable.variable_name)
#
#
#     def test_field_subject(self) :
#
#         #ExStart
#         #ExFor:FieldSubject
#         #ExFor:FieldSubject.text
#         #ExSummary:Shows how to use the SUBJECT field.
#         doc = aw.Document()
#
#         # Set a value for the document's "Subject" built-in property.
#         doc.built_in_document_properties.subject = "My subject"
#
#         # Create a SUBJECT field to display the value of that built-in property.
#         builder = aw.DocumentBuilder(doc)
#         FieldSubject field = (FieldSubject)builder.insert_field(FieldType.field_subject, true)
#         field.update()
#
#         self.assertEqual(" SUBJECT ", field.get_field_code())
#         self.assertEqual("My subject", field.result)
#
#         # If we give the SUBJECT field's Text property value and update it, the field will
#         # overwrite the current value of the "Subject" built-in property with the value of its Text property,
#         # and then display the new value.
#         field.text = "My new subject"
#         field.update()
#
#         self.assertEqual(" SUBJECT  \"My new subject\"", field.get_field_code())
#         self.assertEqual("My new subject", field.result)
#
#         self.assertEqual("My new subject", doc.built_in_document_properties.subject)
#
#         doc.save(aeb.artifacts_dir + "Field.subject.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.subject.docx")
#
#         self.assertEqual("My new subject", doc.built_in_document_properties.subject)
#
#         field = (FieldSubject)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_subject, " SUBJECT  \"My new subject\"", "My new subject", field)
#         self.assertEqual("My new subject", field.text)
#
#
#     def test_field_comments(self) :
#
#         #ExStart
#         #ExFor:FieldComments
#         #ExFor:FieldComments.text
#         #ExSummary:Shows how to use the COMMENTS field.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Set a value for the document's "Comments" built-in property.
#         doc.built_in_document_properties.comments = "My comment."
#
#         # Create a COMMENTS field to display the value of that built-in property.
#         FieldComments field = (FieldComments)builder.insert_field(FieldType.field_comments, true)
#         field.update()
#
#         self.assertEqual(" COMMENTS ", field.get_field_code())
#         self.assertEqual("My comment.", field.result)
#
#         # If we give the COMMENTS field's Text property value and update it, the field will
#         # overwrite the current value of the "Comments" built-in property with the value of its Text property,
#         # and then display the new value.
#         field.text = "My overriding comment."
#         field.update()
#
#         self.assertEqual(" COMMENTS  \"My overriding comment.\"", field.get_field_code())
#         self.assertEqual("My overriding comment.", field.result)
#
#         doc.save(aeb.artifacts_dir + "Field.comments.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.comments.docx")
#
#         self.assertEqual("My overriding comment.", doc.built_in_document_properties.comments)
#
#         field = (FieldComments)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_comments, " COMMENTS  \"My overriding comment.\"", "My overriding comment.", field)
#         self.assertEqual("My overriding comment.", field.text)
#
#
#     def test_field_file_size(self) :
#
#         #ExStart
#         #ExFor:FieldFileSize
#         #ExFor:FieldFileSize.is_in_kilobytes
#         #ExFor:FieldFileSize.is_in_megabytes
#         #ExSummary:Shows how to display the file size of a document with a FILESIZE field.
#         Document doc = new Document(aeb.my_dir + "Document.docx")
#
#         self.assertEqual(18105, doc.built_in_document_properties.bytes)
#
#         builder = aw.DocumentBuilder(doc)
#         builder.move_to_document_end()
#         builder.insert_paragraph()
#
#         # Below are three different units of measure
#         # with which FILESIZE fields can display the document's file size.
#         # 1 -  Bytes:
#         FieldFileSize field = (FieldFileSize)builder.insert_field(FieldType.field_file_size, true)
#         field.update()
#
#         self.assertEqual(" FILESIZE ", field.get_field_code())
#         self.assertEqual("18105", field.result)
#
#         # 2 -  Kilobytes:
#         builder.insert_paragraph()
#         field = (FieldFileSize)builder.insert_field(FieldType.field_file_size, true)
#         field.is_in_kilobytes = true
#         field.update()
#
#         self.assertEqual(" FILESIZE  \\k", field.get_field_code())
#         self.assertEqual("18", field.result)
#
#         # 3 -  Megabytes:
#         builder.insert_paragraph()
#         field = (FieldFileSize)builder.insert_field(FieldType.field_file_size, true)
#         field.is_in_megabytes = true
#         field.update()
#
#         self.assertEqual(" FILESIZE  \\m", field.get_field_code())
#         self.assertEqual("0", field.result)
#
#         # To update the values of these fields while editing in Microsoft Word,
#         # we must first save the changes, and then manually update these fields.
#         doc.save(aeb.artifacts_dir + "Field.filesize.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.filesize.docx")
#
#         field = (FieldFileSize)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_file_size, " FILESIZE ", "18105", field)
#
#         # These fields will need to be updated to produce an accurate result.
#         doc.update_fields()
#
#         field = (FieldFileSize)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_file_size, " FILESIZE  \\k", "13", field)
#         self.assertTrue(field.is_in_kilobytes)
#
#         field = (FieldFileSize)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_file_size, " FILESIZE  \\m", "0", field)
#         self.assertTrue(field.is_in_megabytes)
#
#
#     def test_field_go_to_button(self) :
#
#         #ExStart
#         #ExFor:FieldGoToButton
#         #ExFor:FieldGoToButton.display_text
#         #ExFor:FieldGoToButton.location
#         #ExSummary:Shows to insert a GOTOBUTTON field.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Add a GOTOBUTTON field. When we double-click this field in Microsoft Word,
#         # it will take the text cursor to the bookmark whose name the Location property references.
#         FieldGoToButton field = (FieldGoToButton)builder.insert_field(FieldType.field_go_to_button, true)
#         field.display_text = "My Button"
#         field.location = "MyBookmark"
#
#         self.assertEqual(" GOTOBUTTON  MyBookmark My Button", field.get_field_code())
#
#         # Insert a valid bookmark for the field to reference.
#         builder.insert_break(BreakType.page_break)
#         builder.start_bookmark(field.location)
#         builder.writeln("Bookmark text contents.")
#         builder.end_bookmark(field.location)
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.gotobutton.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.gotobutton.docx")
#         field = (FieldGoToButton)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_go_to_button, " GOTOBUTTON  MyBookmark My Button", string.empty, field)
#         self.assertEqual("My Button", field.display_text)
#         self.assertEqual("MyBookmark", field.location)
#
#
#     [Test]
#     #ExStart
#     #ExFor:FieldFillIn
#     #ExFor:FieldFillIn.default_response
#     #ExFor:FieldFillIn.prompt_once_on_mail_merge
#     #ExFor:FieldFillIn.prompt_text
#     #ExSummary:Shows how to use the FILLIN field to prompt the user for a response.
#     public void FieldFillIn()
#
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Insert a FILLIN field. When we manually update this field in Microsoft Word,
#         # it will prompt us to enter a response. The field will then display the response as text.
#         FieldFillIn field = (FieldFillIn)builder.insert_field(FieldType.field_fill_in, true)
#         field.prompt_text = "Please enter a response:"
#         field.default_response = "A default response."
#
#         # We can also use these fields to ask the user for a unique response for each page
#         # created during a mail merge done using Microsoft Word.
#         field.prompt_once_on_mail_merge = true
#
#         self.assertEqual(" FILLIN  \"Please enter a response:\" \\d \"A default response.\" \\o", field.get_field_code())
#
#         FieldMergeField mergeField = (FieldMergeField)builder.insert_field(FieldType.field_merge_field, true)
#         mergeField.field_name = "MergeField"
#
#         # If we perform a mail merge programmatically, we can use a custom prompt respondent
#         # to automatically edit responses for FILLIN fields that the mail merge encounters.
#         doc.field_options.user_prompt_respondent = new PromptRespondent()
#         doc.mail_merge.execute(new []  "MergeField" , new object[]  "" )
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.fillin.docx")
#         TestFieldFillIn(new Document(aeb.artifacts_dir + "Field.fillin.docx")) #ExSKip
#
#
#     # <summary>
#     # Prepends a line to the default response of every FILLIN field during a mail merge.
#     # </summary>
#     private class PromptRespondent : IFieldUserPromptRespondent
#
#         public string Respond(string promptText, string defaultResponse)
#
#             return "Response modified by PromptRespondent. " + defaultResponse
#
#
#     #ExEnd
#
#     private void TestFieldFillIn(Document doc)
#
#         doc = DocumentHelper.save_open(doc)
#
#         self.assertEqual(1, doc.range.fields.count)
#
#         FieldFillIn field = (FieldFillIn)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_fill_in, " FILLIN  \"Please enter a response:\" \\d \"A default response.\" \\o",
#             "Response modified by PromptRespondent. A default response.", field)
#         self.assertEqual("Please enter a response:", field.prompt_text)
#         self.assertEqual("A default response.", field.default_response)
#         self.assertTrue(field.prompt_once_on_mail_merge)
#
#
#     def test_field_info(self) :
#
#         #ExStart
#         #ExFor:FieldInfo
#         #ExFor:FieldInfo.info_type
#         #ExFor:FieldInfo.new_value
#         #ExSummary:Shows how to work with INFO fields.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Set a value for the "Comments" built-in property and then insert an INFO field to display that property's value.
#         doc.built_in_document_properties.comments = "My comment"
#         FieldInfo field = (FieldInfo)builder.insert_field(FieldType.field_info, true)
#         field.info_type = "Comments"
#         field.update()
#
#         self.assertEqual(" INFO  Comments", field.get_field_code())
#         self.assertEqual("My comment", field.result)
#
#         builder.writeln()
#
#         # Setting a value for the field's NewValue property and updating
#         # the field will also overwrite the corresponding built-in property with the new value.
#         field = (FieldInfo)builder.insert_field(FieldType.field_info, true)
#         field.info_type = "Comments"
#         field.new_value = "New comment"
#         field.update()
#
#         self.assertEqual(" INFO  Comments \"New comment\"", field.get_field_code())
#         self.assertEqual("New comment", field.result)
#         self.assertEqual("New comment", doc.built_in_document_properties.comments)
#
#         doc.save(aeb.artifacts_dir + "Field.info.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.info.docx")
#
#         self.assertEqual("New comment", doc.built_in_document_properties.comments)
#
#         field = (FieldInfo)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_info, " INFO  Comments", "My comment", field)
#         self.assertEqual("Comments", field.info_type)
#
#         field = (FieldInfo)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_info, " INFO  Comments \"New comment\"", "New comment", field)
#         self.assertEqual("Comments", field.info_type)
#         self.assertEqual("New comment", field.new_value)
#
#
#     def test_field_macro_button(self) :
#
#         #ExStart
#         #ExFor:Document.has_macros
#         #ExFor:FieldMacroButton
#         #ExFor:FieldMacroButton.display_text
#         #ExFor:FieldMacroButton.macro_name
#         #ExSummary:Shows how to use MACROBUTTON fields to allow us to run a document's macros by clicking.
#         Document doc = new Document(aeb.my_dir + "Macro.docm")
#         builder = aw.DocumentBuilder(doc)
#
#         self.assertTrue(doc.has_macros)
#
#         # Insert a MACROBUTTON field, and reference one of the document's macros by name in the MacroName property.
#         FieldMacroButton field = (FieldMacroButton)builder.insert_field(FieldType.field_macro_button, true)
#         field.macro_name = "MyMacro"
#         field.display_text = "Double click to run macro: " + field.macro_name
#
#         self.assertEqual(" MACROBUTTON  MyMacro Double click to run macro: MyMacro", field.get_field_code())
#
#         # Use the property to reference "ViewZoom200", a macro that ships with Microsoft Word.
#         # We can find all other macros via View -> Macros (dropdown) -> View Macros.
#         # In that menu, select "Word Commands" from the "Macros in:" drop down.
#         # If our document contains a custom macro with the same name as a stock macro,
#         # our macro will be the one that the MACROBUTTON field runs.
#         builder.insert_paragraph()
#         field = (FieldMacroButton)builder.insert_field(FieldType.field_macro_button, true)
#         field.macro_name = "ViewZoom200"
#         field.display_text = "Run " + field.macro_name
#
#         self.assertEqual(" MACROBUTTON  ViewZoom200 Run ViewZoom200", field.get_field_code())
#
#         # Save the document as a macro-enabled document type.
#         doc.save(aeb.artifacts_dir + "Field.macrobutton.docm")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.macrobutton.docm")
#
#         field = (FieldMacroButton)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_macro_button, " MACROBUTTON  MyMacro Double click to run macro: MyMacro", string.empty, field)
#         self.assertEqual("MyMacro", field.macro_name)
#         self.assertEqual("Double click to run macro: MyMacro", field.display_text)
#
#         field = (FieldMacroButton)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_macro_button, " MACROBUTTON  ViewZoom200 Run ViewZoom200", string.empty, field)
#         self.assertEqual("ViewZoom200", field.macro_name)
#         self.assertEqual("Run ViewZoom200", field.display_text)
#
#
#     def test_field_keywords(self) :
#
#         #ExStart
#         #ExFor:FieldKeywords
#         #ExFor:FieldKeywords.text
#         #ExSummary:Shows to insert a KEYWORDS field.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Add some keywords, also referred to as "tags" in File Explorer.
#         doc.built_in_document_properties.keywords = "Keyword1, Keyword2"
#
#         # The KEYWORDS field displays the value of this property.
#         FieldKeywords field = (FieldKeywords)builder.insert_field(FieldType.field_keyword, true)
#         field.update()
#
#         self.assertEqual(" KEYWORDS ", field.get_field_code())
#         self.assertEqual("Keyword1, Keyword2", field.result)
#
#         # Setting a value for the field's Text property,
#         # and then updating the field will also overwrite the corresponding built-in property with the new value.
#         field.text = "OverridingKeyword"
#         field.update()
#
#         self.assertEqual(" KEYWORDS  OverridingKeyword", field.get_field_code())
#         self.assertEqual("OverridingKeyword", field.result)
#         self.assertEqual("OverridingKeyword", doc.built_in_document_properties.keywords)
#
#         doc.save(aeb.artifacts_dir + "Field.keywords.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.keywords.docx")
#
#         self.assertEqual("OverridingKeyword", doc.built_in_document_properties.keywords)
#
#         field = (FieldKeywords)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_keyword, " KEYWORDS  OverridingKeyword", "OverridingKeyword", field)
#         self.assertEqual("OverridingKeyword", field.text)
#
#
#     def test_field_num(self) :
#
#         #ExStart
#         #ExFor:FieldPage
#         #ExFor:FieldNumChars
#         #ExFor:FieldNumPages
#         #ExFor:FieldNumWords
#         #ExSummary:Shows how to use NUMCHARS, NUMWORDS, NUMPAGES and PAGE fields to track the size of our documents.
#         Document doc = new Document(aeb.my_dir + "Paragraphs.docx")
#         builder = aw.DocumentBuilder(doc)
#
#         builder.move_to_header_footer(HeaderFooterType.footer_primary)
#         builder.paragraph_format.alignment = ParagraphAlignment.center
#
#         # Below are three types of fields that we can use to track the size of our documents.
#         # 1 -  Track the character count with a NUMCHARS field:
#         FieldNumChars fieldNumChars = (FieldNumChars)builder.insert_field(FieldType.field_num_chars, true)
#         builder.writeln(" characters")
#
#         # 2 -  Track the word count with a NUMWORDS field:
#         FieldNumWords fieldNumWords = (FieldNumWords)builder.insert_field(FieldType.field_num_words, true)
#         builder.writeln(" words")
#
#         # 3 -  Use both PAGE and NUMPAGES fields to display what page the field is on,
#         # and the total number of pages in the document:
#         builder.paragraph_format.alignment = ParagraphAlignment.right
#         builder.write("Page ")
#         FieldPage fieldPage = (FieldPage)builder.insert_field(FieldType.field_page, true)
#         builder.write(" of ")
#         FieldNumPages fieldNumPages = (FieldNumPages)builder.insert_field(FieldType.field_num_pages, true)
#
#         self.assertEqual(" NUMCHARS ", fieldNumChars.get_field_code())
#         self.assertEqual(" NUMWORDS ", fieldNumWords.get_field_code())
#         self.assertEqual(" NUMPAGES ", fieldNumPages.get_field_code())
#         self.assertEqual(" PAGE ", fieldPage.get_field_code())
#
#         # These fields will not maintain accurate values in real time
#         # while we edit the document programmatically using Aspose.words, or in Microsoft Word.
#         # We need to update them every we need to see an up-to-date value.
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.numchars.numwords.numpages.page.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.numchars.numwords.numpages.page.docx")
#
#         TestUtil.verify_field(FieldType.field_num_chars, " NUMCHARS ", "6009", doc.range.fields[0])
#         TestUtil.verify_field(FieldType.field_num_words, " NUMWORDS ", "1054", doc.range.fields[1])
#
#         TestUtil.verify_field(FieldType.field_page, " PAGE ", "6", doc.range.fields[2])
#         TestUtil.verify_field(FieldType.field_num_pages, " NUMPAGES ", "6", doc.range.fields[3])
#
#
#     def test_field_print(self) :
#
#         #ExStart
#         #ExFor:FieldPrint
#         #ExFor:FieldPrint.post_script_group
#         #ExFor:FieldPrint.printer_instructions
#         #ExSummary:Shows to insert a PRINT field.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         builder.write("My paragraph")
#
#         # The PRINT field can send instructions to the printer.
#         FieldPrint field = (FieldPrint)builder.insert_field(FieldType.field_print, true)
#
#         # Set the area for the printer to perform instructions over.
#         # In this case, it will be the paragraph that contains our PRINT field.
#         field.post_script_group = "para"
#
#         # When we use a printer that supports PostScript to print our document,
#         # this command will turn the entire area that we specified in "field.post_script_group" white.
#         field.printer_instructions = "erasepage"
#
#         self.assertEqual(" PRINT  erasepage \\p para", field.get_field_code())
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.print.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.print.docx")
#
#         field = (FieldPrint)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_print, " PRINT  erasepage \\p para", string.empty, field)
#         self.assertEqual("para", field.post_script_group)
#         self.assertEqual("erasepage", field.printer_instructions)
#
#
#     def test_field_print_date(self) :
#
#         #ExStart
#         #ExFor:FieldPrintDate
#         #ExFor:FieldPrintDate.use_lunar_calendar
#         #ExFor:FieldPrintDate.use_saka_era_calendar
#         #ExFor:FieldPrintDate.use_um_al_qura_calendar
#         #ExSummary:Shows read PRINTDATE fields.
#         Document doc = new Document(aeb.my_dir + "Field sample - PRINTDATE.docx")
#
#         # When a document is printed by a printer or printed as a PDF (but not exported to PDF),
#         # PRINTDATE fields will display the print operation's date/time.
#         # If no printing has taken place, these fields will display "0/0/0000".
#         FieldPrintDate field = (FieldPrintDate)doc.range.fields[0]
#
#         self.assertEqual("3/25/2020 12:00:00 AM", field.result)
#         self.assertEqual(" PRINTDATE ", field.get_field_code())
#
#         # Below are three different calendar types according to which the PRINTDATE field
#         # can display the date and time of the last printing operation.
#         # 1 -  Islamic Lunar Calendar:
#         field = (FieldPrintDate)doc.range.fields[1]
#
#         self.assertTrue(field.use_lunar_calendar)
#         self.assertEqual("8/1/1441 12:00:00 AM", field.result)
#         self.assertEqual(" PRINTDATE  \\h", field.get_field_code())
#
#         field = (FieldPrintDate)doc.range.fields[2]
#
#         # 2 -  Umm al-Qura calendar:
#         self.assertTrue(field.use_um_al_qura_calendar)
#         self.assertEqual("8/1/1441 12:00:00 AM", field.result)
#         self.assertEqual(" PRINTDATE  \\u", field.get_field_code())
#
#         field = (FieldPrintDate)doc.range.fields[3]
#
#         # 3 -  Indian National Calendar:
#         self.assertTrue(field.use_saka_era_calendar)
#         self.assertEqual("1/5/1942 12:00:00 AM", field.result)
#         self.assertEqual(" PRINTDATE  \\s", field.get_field_code())
#         #ExEnd
#
#
#     def test_field_quote(self) :
#
#         #ExStart
#         #ExFor:FieldQuote
#         #ExFor:FieldQuote.text
#         #ExFor:Document.update_fields
#         #ExSummary:Shows to use the QUOTE field.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Insert a QUOTE field, which will display the value of its Text property.
#         FieldQuote field = (FieldQuote)builder.insert_field(FieldType.field_quote, true)
#         field.text = "\"Quoted text\""
#
#         self.assertEqual(" QUOTE  \"\\\"Quoted text\\\"\"", field.get_field_code())
#
#         # Insert a QUOTE field and nest a DATE field inside it.
#         # DATE fields update their value to the current date every time we open the document using Microsoft Word.
#         # Nesting the DATE field inside the QUOTE field like this will freeze its value
#         # to the date when we created the document.
#         builder.write("\nDocument creation date: ")
#         field = (FieldQuote)builder.insert_field(FieldType.field_quote, true)
#         builder.move_to(field.separator)
#         builder.insert_field(FieldType.field_date, true)
#
#         self.assertEqual(" QUOTE \u0013 DATE \u0014" + DateTime.now.date.to_short_date_string() + "\u0015", field.get_field_code())
#
#         # Update all the fields to display their correct results.
#         doc.update_fields()
#
#         self.assertEqual("\"Quoted text\"", doc.range.fields[0].result)
#
#         doc.save(aeb.artifacts_dir + "Field.quote.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.quote.docx")
#
#         TestUtil.verify_field(FieldType.field_quote, " QUOTE  \"\\\"Quoted text\\\"\"", "\"Quoted text\"", doc.range.fields[0])
#
#         TestUtil.verify_field(FieldType.field_quote, " QUOTE \u0013 DATE \u0014" + DateTime.now.date.to_short_date_string() + "\u0015",
#             DateTime.now.date.to_short_date_string(), doc.range.fields[1])
#
#
#
#     #ExStart
#     #ExFor:FieldNext
#     #ExFor:FieldNextIf
#     #ExFor:FieldNextIf.comparison_operator
#     #ExFor:FieldNextIf.left_expression
#     #ExFor:FieldNextIf.right_expression
#     #ExSummary:Shows how to use NEXT/NEXTIF fields to merge multiple rows into one page during a mail merge.
#     def test_field_next(self) :
#
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Create a data source for our mail merge with 3 rows.
#         # A mail merge that uses this table would normally create a 3-page document.
#         DataTable table = new DataTable("Employees")
#         table.columns.add("Courtesy Title")
#         table.columns.add("First Name")
#         table.columns.add("Last Name")
#         table.rows.add("Mr.", "John", "Doe")
#         table.rows.add("Mrs.", "Jane", "Cardholder")
#         table.rows.add("Mr.", "Joe", "Bloggs")
#
#         InsertMergeFields(builder, "First row: ")
#
#         # If we have multiple merge fields with the same FieldName,
#         # they will receive data from the same row of the data source and display the same value after the merge.
#         # A NEXT field tells the mail merge instantly to move down one row,
#         # which means any MERGEFIELDs that follow the NEXT field will receive data from the next row.
#         # Make sure never to try to skip to the next row while already on the last row.
#         FieldNext fieldNext = (FieldNext)builder.insert_field(FieldType.field_next, true)
#
#         self.assertEqual(" NEXT ", fieldNext.get_field_code())
#
#         # After the merge, the data source values that these MERGEFIELDs accept
#         # will end up on the same page as the MERGEFIELDs above.
#         InsertMergeFields(builder, "Second row: ")
#
#         # A NEXTIF field has the same function as a NEXT field,
#         # but it skips to the next row only if a statement constructed by the following 3 properties is true.
#         FieldNextIf fieldNextIf = (FieldNextIf)builder.insert_field(FieldType.field_next_if, true)
#         fieldNextIf.left_expression = "5"
#         fieldNextIf.right_expression = "2 + 3"
#         fieldNextIf.comparison_operator = "="
#
#         self.assertEqual(" NEXTIF  5 = \"2 + 3\"", fieldNextIf.get_field_code())
#
#         # If the comparison asserted by the above field is correct,
#         # the following 3 merge fields will take data from the third row.
#         # Otherwise, these fields will take data from row 2 again.
#         InsertMergeFields(builder, "Third row: ")
#
#         doc.mail_merge.execute(table)
#
#         # Our data source has 3 rows, and we skipped rows twice.
#         # Our output document will have 1 page with data from all 3 rows.
#         doc.save(aeb.artifacts_dir + "Field.next.nextif.docx")
#         TestFieldNext(doc) #ExSKip
#
#
#     # <summary>
#     # Uses a document builder to insert MERGEFIELDs for a data source that contains columns named "Courtesy Title", "First Name" and "Last Name".
#     # </summary>
#     public void InsertMergeFields(DocumentBuilder builder, string firstFieldTextBefore)
#
#         InsertMergeField(builder, "Courtesy Title", firstFieldTextBefore, " ")
#         InsertMergeField(builder, "First Name", null, " ")
#         InsertMergeField(builder, "Last Name", null, null)
#         builder.insert_paragraph()
#
#
#     # <summary>
#     # Uses a document builder to insert a MERRGEFIELD with specified properties.
#     # </summary>
#     public void InsertMergeField(DocumentBuilder builder, string fieldName, string textBefore, string textAfter)
#
#         FieldMergeField field = (FieldMergeField) builder.insert_field(FieldType.field_merge_field, true)
#         field.field_name = fieldName
#         field.text_before = textBefore
#         field.text_after = textAfter
#
#     #ExEnd
#
#     private void TestFieldNext(Document doc)
#
#         doc = DocumentHelper.save_open(doc)
#
#         self.assertEqual(0, doc.range.fields.count)
#         self.assertEqual("First row: Mr. John Doe\r" +
#                         "Second row: Mrs. Jane Cardholder\r" +
#                         "Third row: Mr. Joe Bloggs\r\f", doc.get_text())
#
#
#     #ExStart
#     #ExFor:FieldNoteRef
#     #ExFor:FieldNoteRef.bookmark_name
#     #ExFor:FieldNoteRef.insert_hyperlink
#     #ExFor:FieldNoteRef.insert_reference_mark
#     #ExFor:FieldNoteRef.insert_relative_position
#     #ExSummary:Shows to insert NOTEREF fields, and modify their appearance.
#     [Test] #ExSkip
#     [Ignore("WORDSNET-17845")] #ExSkip
#     public void FieldNoteRef()
#
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Create a bookmark with a footnote that the NOTEREF field will reference.
#         InsertBookmarkWithFootnote(builder, "MyBookmark1", "Contents of MyBookmark1", "Footnote from MyBookmark1")
#
#         # This NOTEREF field will display the number of the footnote inside the referenced bookmark.
#         # Setting the InsertHyperlink property lets us jump to the bookmark by Ctrl + clicking the field in Microsoft Word.
#         self.assertEqual(" NOTEREF  MyBookmark2 \\h",
#             InsertFieldNoteRef(builder, "MyBookmark2", true, false, false, "Hyperlink to Bookmark2, with footnote number ").get_field_code())
#
#         # When using the \p flag, after the footnote number, the field also displays the bookmark's position relative to the field.
#         # Bookmark1 is above this field and contains footnote number 1, so the result will be "1 above" on update.
#         self.assertEqual(" NOTEREF  MyBookmark1 \\h \\p",
#             InsertFieldNoteRef(builder, "MyBookmark1", true, true, false, "Bookmark1, with footnote number ").get_field_code())
#
#         # Bookmark2 is below this field and contains footnote number 2, so the field will display "2 below".
#         # The \f flag makes the number 2 appear in the same format as the footnote number label in the actual text.
#         self.assertEqual(" NOTEREF  MyBookmark2 \\h \\p \\f",
#             InsertFieldNoteRef(builder, "MyBookmark2", true, true, true, "Bookmark2, with footnote number ").get_field_code())
#
#         builder.insert_break(BreakType.page_break)
#         InsertBookmarkWithFootnote(builder, "MyBookmark2", "Contents of MyBookmark2", "Footnote from MyBookmark2")
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.noteref.docx")
#         TestNoteRef(new Document(aeb.artifacts_dir + "Field.noteref.docx")) #ExSkip
#
#
#     # <summary>
#     # Uses a document builder to insert a NOTEREF field with specified properties.
#     # </summary>
#     private static FieldNoteRef InsertFieldNoteRef(DocumentBuilder builder, string bookmarkName, bool insertHyperlink, bool insertRelativePosition, bool insertReferenceMark, string textBefore)
#
#         builder.write(textBefore)
#
#         FieldNoteRef field = (FieldNoteRef)builder.insert_field(FieldType.field_note_ref, true)
#         field.bookmark_name = bookmarkName
#         field.insert_hyperlink = insertHyperlink
#         field.insert_relative_position = insertRelativePosition
#         field.insert_reference_mark = insertReferenceMark
#         builder.writeln()
#
#         return field
#
#
#     # <summary>
#     # Uses a document builder to insert a named bookmark with a footnote at the end.
#     # </summary>
#     private static void InsertBookmarkWithFootnote(DocumentBuilder builder, string bookmarkName, string bookmarkText, string footnoteText)
#
#         builder.start_bookmark(bookmarkName)
#         builder.write(bookmarkText)
#         builder.insert_footnote(FootnoteType.footnote, footnoteText)
#         builder.end_bookmark(bookmarkName)
#         builder.writeln()
#
#     #ExEnd
#
#     private void TestNoteRef(Document doc)
#
#         FieldNoteRef field = (FieldNoteRef)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_note_ref, " NOTEREF  MyBookmark2 \\h", "2", field)
#         self.assertEqual("MyBookmark2", field.bookmark_name)
#         self.assertTrue(field.insert_hyperlink)
#         self.assertFalse(field.insert_relative_position)
#         self.assertFalse(field.insert_reference_mark)
#
#         field = (FieldNoteRef)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_note_ref, " NOTEREF  MyBookmark1 \\h \\p", "1 above", field)
#         self.assertEqual("MyBookmark1", field.bookmark_name)
#         self.assertTrue(field.insert_hyperlink)
#         self.assertTrue(field.insert_relative_position)
#         self.assertFalse(field.insert_reference_mark)
#
#         field = (FieldNoteRef)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_note_ref, " NOTEREF  MyBookmark2 \\h \\p \\f", "2 below", field)
#         self.assertEqual("MyBookmark2", field.bookmark_name)
#         self.assertTrue(field.insert_hyperlink)
#         self.assertTrue(field.insert_relative_position)
#         self.assertTrue(field.insert_reference_mark)
#
#
#     [Test]
#     [Ignore("WORDSNET-17845")]
#     public void FootnoteRef()
#
#         #ExStart
#         #ExFor:FieldFootnoteRef
#         #ExSummary:Shows how to cross-reference footnotes with the FOOTNOTEREF field.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         builder.start_bookmark("CrossRefBookmark")
#         builder.write("Hello world!")
#         builder.insert_footnote(FootnoteType.footnote, "Cross referenced footnote.")
#         builder.end_bookmark("CrossRefBookmark")
#         builder.insert_paragraph()
#
#         # Insert a FOOTNOTEREF field, which lets us reference a footnote more than once while re-using the same footnote marker.
#         builder.write("CrossReference: ")
#         FieldFootnoteRef field = (FieldFootnoteRef) builder.insert_field(FieldType.field_footnote_ref, true)
#
#         # Reference the bookmark that we have created with the FOOTNOTEREF field. That bookmark contains a footnote marker
#         # belonging to the footnote we inserted. The field will display that footnote marker.
#         builder.move_to(field.separator)
#         builder.write("CrossRefBookmark")
#
#         self.assertEqual(" FOOTNOTEREF CrossRefBookmark", field.get_field_code())
#
#         doc.update_fields()
#
#         # This field works only in older versions of Microsoft Word.
#         doc.save(aeb.artifacts_dir + "Field.footnoteref.doc")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.footnoteref.doc")
#         field = (FieldFootnoteRef)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_footnote_ref, " FOOTNOTEREF CrossRefBookmark", "1", field)
#         TestUtil.verify_footnote(FootnoteType.footnote, true, string.empty, "Cross referenced footnote.",
#             (Footnote)doc.get_child(NodeType.footnote, 0, true))
#
#
#     #ExStart
#     #ExFor:FieldPageRef
#     #ExFor:FieldPageRef.bookmark_name
#     #ExFor:FieldPageRef.insert_hyperlink
#     #ExFor:FieldPageRef.insert_relative_position
#     #ExSummary:Shows to insert PAGEREF fields to display the relative location of bookmarks.
#     [Test] #ExSkip
#     [Ignore("WORDSNET-17836")] #ExSkip
#     public void FieldPageRef()
#
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         InsertAndNameBookmark(builder, "MyBookmark1")
#
#         # Insert a PAGEREF field that displays what page a bookmark is on.
#         # Set the InsertHyperlink flag to make the field also function as a clickable link to the bookmark.
#         self.assertEqual(" PAGEREF  MyBookmark3 \\h",
#             InsertFieldPageRef(builder, "MyBookmark3", true, false, "Hyperlink to Bookmark3, on page: ").get_field_code())
#
#         # We can use the \p flag to get the PAGEREF field to display
#         # the bookmark's position relative to the position of the field.
#         # Bookmark1 is on the same page and above this field, so this field's displayed result will be "above".
#         self.assertEqual(" PAGEREF  MyBookmark1 \\h \\p",
#             InsertFieldPageRef(builder, "MyBookmark1", true, true, "Bookmark1 is ").get_field_code())
#
#         # Bookmark2 will be on the same page and below this field, so this field's displayed result will be "below".
#         self.assertEqual(" PAGEREF  MyBookmark2 \\h \\p",
#             InsertFieldPageRef(builder, "MyBookmark2", true, true, "Bookmark2 is ").get_field_code())
#
#         # Bookmark3 will be on a different page, so the field will display "on page 2".
#         self.assertEqual(" PAGEREF  MyBookmark3 \\h \\p",
#             InsertFieldPageRef(builder, "MyBookmark3", true, true, "Bookmark3 is ").get_field_code())
#
#         InsertAndNameBookmark(builder, "MyBookmark2")
#         builder.insert_break(BreakType.page_break)
#         InsertAndNameBookmark(builder, "MyBookmark3")
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.pageref.docx")
#         TestPageRef(new Document(aeb.artifacts_dir + "Field.pageref.docx")) #ExSkip
#
#
#     # <summary>
#     # Uses a document builder to insert a PAGEREF field and sets its properties.
#     # </summary>
#     private static FieldPageRef InsertFieldPageRef(DocumentBuilder builder, string bookmarkName, bool insertHyperlink, bool insertRelativePosition, string textBefore)
#
#         builder.write(textBefore)
#
#         FieldPageRef field = (FieldPageRef)builder.insert_field(FieldType.field_page_ref, true)
#         field.bookmark_name = bookmarkName
#         field.insert_hyperlink = insertHyperlink
#         field.insert_relative_position = insertRelativePosition
#         builder.writeln()
#
#         return field
#
#
#     # <summary>
#     # Uses a document builder to insert a named bookmark.
#     # </summary>
#     private static void InsertAndNameBookmark(DocumentBuilder builder, string bookmarkName)
#
#         builder.start_bookmark(bookmarkName)
#         builder.writeln($"Contents of bookmark \"bookmarkName\".")
#         builder.end_bookmark(bookmarkName)
#
#     #ExEnd
#
#     private void TestPageRef(Document doc)
#
#         FieldPageRef field = (FieldPageRef)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_page_ref, " PAGEREF  MyBookmark3 \\h", "2", field)
#         self.assertEqual("MyBookmark3", field.bookmark_name)
#         self.assertTrue(field.insert_hyperlink)
#         self.assertFalse(field.insert_relative_position)
#
#         field = (FieldPageRef)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_page_ref, " PAGEREF  MyBookmark1 \\h \\p", "above", field)
#         self.assertEqual("MyBookmark1", field.bookmark_name)
#         self.assertTrue(field.insert_hyperlink)
#         self.assertTrue(field.insert_relative_position)
#
#         field = (FieldPageRef)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_page_ref, " PAGEREF  MyBookmark2 \\h \\p", "below", field)
#         self.assertEqual("MyBookmark2", field.bookmark_name)
#         self.assertTrue(field.insert_hyperlink)
#         self.assertTrue(field.insert_relative_position)
#
#         field = (FieldPageRef)doc.range.fields[3]
#
#         TestUtil.verify_field(FieldType.field_page_ref, " PAGEREF  MyBookmark3 \\h \\p", "on page 2", field)
#         self.assertEqual("MyBookmark3", field.bookmark_name)
#         self.assertTrue(field.insert_hyperlink)
#         self.assertTrue(field.insert_relative_position)
#
#
#     #ExStart
#     #ExFor:FieldRef
#     #ExFor:FieldRef.bookmark_name
#     #ExFor:FieldRef.include_note_or_comment
#     #ExFor:FieldRef.insert_hyperlink
#     #ExFor:FieldRef.insert_paragraph_number
#     #ExFor:FieldRef.insert_paragraph_number_in_full_context
#     #ExFor:FieldRef.insert_paragraph_number_in_relative_context
#     #ExFor:FieldRef.insert_relative_position
#     #ExFor:FieldRef.number_separator
#     #ExFor:FieldRef.suppress_non_delimiters
#     #ExSummary:Shows how to insert REF fields to reference bookmarks.
#     [Test] #ExSkip
#     [Ignore("WORDSNET-18067")] #ExSkip
#     public void FieldRef()
#
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         builder.start_bookmark("MyBookmark")
#         builder.insert_footnote(FootnoteType.footnote, "MyBookmark footnote #1")
#         builder.write("Text that will appear in REF field")
#         builder.insert_footnote(FootnoteType.footnote, "MyBookmark footnote #2")
#         builder.end_bookmark("MyBookmark")
#         builder.move_to_document_start()
#
#         # We will apply a custom list format, where the amount of angle brackets indicates the list level we are currently at.
#         builder.list_format.apply_number_default()
#         builder.list_format.list_level.number_format = "> \x0000"
#
#         # Insert a REF field that will contain the text within our bookmark, act as a hyperlink, and clone the bookmark's footnotes.
#         FieldRef field = InsertFieldRef(builder, "MyBookmark", "", "\n")
#         field.include_note_or_comment = true
#         field.insert_hyperlink = true
#
#         self.assertEqual(" REF  MyBookmark \\f \\h", field.get_field_code())
#
#         # Insert a REF field, and display whether the referenced bookmark is above or below it.
#         field = InsertFieldRef(builder, "MyBookmark", "The referenced paragraph is ", " this field.\n")
#         field.insert_relative_position = true
#
#         self.assertEqual(" REF  MyBookmark \\p", field.get_field_code())
#
#         # Display the list number of the bookmark as it appears in the document.
#         field = InsertFieldRef(builder, "MyBookmark", "The bookmark's paragraph number is ", "\n")
#         field.insert_paragraph_number = true
#
#         self.assertEqual(" REF  MyBookmark \\n", field.get_field_code())
#
#         # Display the bookmark's list number, but with non-delimiter characters, such as the angle brackets, omitted.
#         field = InsertFieldRef(builder, "MyBookmark", "The bookmark's paragraph number, non-delimiters suppressed, is ", "\n")
#         field.insert_paragraph_number = true
#         field.suppress_non_delimiters = true
#
#         self.assertEqual(" REF  MyBookmark \\n \\t", field.get_field_code())
#
#         # Move down one list level.
#         builder.list_format.list_level_number++
#         builder.list_format.list_level.number_format = ">> \x0001"
#
#         # Display the list number of the bookmark and the numbers of all the list levels above it.
#         field = InsertFieldRef(builder, "MyBookmark", "The bookmark's full context paragraph number is ", "\n")
#         field.insert_paragraph_number_in_full_context = true
#
#         self.assertEqual(" REF  MyBookmark \\w", field.get_field_code())
#
#         builder.insert_break(BreakType.page_break)
#
#         # Display the list level numbers between this REF field, and the bookmark that it is referencing.
#         field = InsertFieldRef(builder, "MyBookmark", "The bookmark's relative paragraph number is ", "\n")
#         field.insert_paragraph_number_in_relative_context = true
#
#         self.assertEqual(" REF  MyBookmark \\r", field.get_field_code())
#
#         # At the end of the document, the bookmark will show up as a list item here.
#         builder.writeln("List level above bookmark")
#         builder.list_format.list_level_number++
#         builder.list_format.list_level.number_format = ">>> \x0002"
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.ref.docx")
#         TestFieldRef(new Document(aeb.artifacts_dir + "Field.ref.docx")) #ExSkip
#
#
#     # <summary>
#     # Get the document builder to insert a REF field, reference a bookmark with it, and add text before and after it.
#     # </summary>
#     private static FieldRef InsertFieldRef(DocumentBuilder builder, string bookmarkName, string textBefore, string textAfter)
#
#         builder.write(textBefore)
#         FieldRef field = (FieldRef)builder.insert_field(FieldType.field_ref, true)
#         field.bookmark_name = bookmarkName
#         builder.write(textAfter)
#         return field
#
#     #ExEnd
#
#     private void TestFieldRef(Document doc)
#
#         TestUtil.verify_footnote(FootnoteType.footnote, true, string.empty, "MyBookmark footnote #1",
#             (Footnote)doc.get_child(NodeType.footnote, 0, true))
#         TestUtil.verify_footnote(FootnoteType.footnote, true, string.empty, "MyBookmark footnote #2",
#             (Footnote)doc.get_child(NodeType.footnote, 0, true))
#
#         FieldRef field = (FieldRef)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_ref, " REF  MyBookmark \\f \\h",
#             "\u0002 MyBookmark footnote #1\r" +
#             "Text that will appear in REF field\u0002 MyBookmark footnote #2\r", field)
#         self.assertEqual("MyBookmark", field.bookmark_name)
#         self.assertTrue(field.include_note_or_comment)
#         self.assertTrue(field.insert_hyperlink)
#
#         field = (FieldRef)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_ref, " REF  MyBookmark \\p", "below", field)
#         self.assertEqual("MyBookmark", field.bookmark_name)
#         self.assertTrue(field.insert_relative_position)
#
#         field = (FieldRef)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_ref, " REF  MyBookmark \\n", ">>> i", field)
#         self.assertEqual("MyBookmark", field.bookmark_name)
#         self.assertTrue(field.insert_paragraph_number)
#         self.assertEqual(" REF  MyBookmark \\n", field.get_field_code())
#         self.assertEqual(">>> i", field.result)
#
#         field = (FieldRef)doc.range.fields[3]
#
#         TestUtil.verify_field(FieldType.field_ref, " REF  MyBookmark \\n \\t", "i", field)
#         self.assertEqual("MyBookmark", field.bookmark_name)
#         self.assertTrue(field.insert_paragraph_number)
#         self.assertTrue(field.suppress_non_delimiters)
#
#         field = (FieldRef)doc.range.fields[4]
#
#         TestUtil.verify_field(FieldType.field_ref, " REF  MyBookmark \\w", "> 4>> c>>> i", field)
#         self.assertEqual("MyBookmark", field.bookmark_name)
#         self.assertTrue(field.insert_paragraph_number_in_full_context)
#
#         field = (FieldRef)doc.range.fields[5]
#
#         TestUtil.verify_field(FieldType.field_ref, " REF  MyBookmark \\r", ">> c>>> i", field)
#         self.assertEqual("MyBookmark", field.bookmark_name)
#         self.assertTrue(field.insert_paragraph_number_in_relative_context)
#
#
#     [Test]
#     [Ignore("WORDSNET-18068")]
#     public void FieldRD()
#
#         #ExStart
#         #ExFor:FieldRD
#         #ExFor:FieldRD.file_name
#         #ExFor:FieldRD.is_path_relative
#         #ExSummary:Shows to use the RD field to create a table of contents entries from headings in other documents.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Use a document builder to insert a table of contents,
#         # and then add one entry for the table of contents on the following page.
#         builder.insert_field(FieldType.field_toc, true)
#         builder.insert_break(BreakType.page_break)
#         builder.current_paragraph.paragraph_format.style_name = "Heading 1"
#         builder.writeln("TOC entry from within this document")
#
#         # Insert an RD field, which references another local file system document in its FileName property.
#         # The TOC will also now accept all headings from the referenced document as entries for its table.
#         FieldRD field = (FieldRD)builder.insert_field(FieldType.field_ref_doc, true)
#         field.file_name = "ReferencedDocument.docx"
#         field.is_path_relative = true
#
#         self.assertEqual(" RD  ReferencedDocument.docx \\f", field.get_field_code())
#
#         # Create the document that the RD field is referencing and insert a heading.
#         # This heading will show up as an entry in the TOC field in our first document.
#         Document referencedDoc = new Document()
#         DocumentBuilder refDocBuilder = new DocumentBuilder(referencedDoc)
#         refDocBuilder.current_paragraph.paragraph_format.style_name = "Heading 1"
#         refDocBuilder.writeln("TOC entry from referenced document")
#         referencedDoc.save(aeb.artifacts_dir + "ReferencedDocument.docx")
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.rd.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.rd.docx")
#
#         FieldToc fieldToc = (FieldToc)doc.range.fields[0]
#
#         self.assertEqual("TOC entry from within this document\t\u0013 PAGEREF _Toc36149519 \\h \u00142\u0015\r" +
#                         "TOC entry from referenced document\t1\r", fieldToc.result)
#
#         FieldPageRef fieldPageRef = (FieldPageRef)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_page_ref, " PAGEREF _Toc36149519 \\h ", "2", fieldPageRef)
#
#         field = (FieldRD)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_ref_doc, " RD  ReferencedDocument.docx \\f", string.empty, field)
#         self.assertEqual("ReferencedDocument.docx", field.file_name)
#         self.assertTrue(field.is_path_relative)
#
#
#     def test_skip_if(self) :
#
#         #ExStart
#         #ExFor:FieldSkipIf
#         #ExFor:FieldSkipIf.comparison_operator
#         #ExFor:FieldSkipIf.left_expression
#         #ExFor:FieldSkipIf.right_expression
#         #ExSummary:Shows how to skip pages in a mail merge using the SKIPIF field.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Insert a SKIPIF field. If the current row of a mail merge operation fulfills the condition
#         # which the expressions of this field state, then the mail merge operation aborts the current row,
#         # discards the current merge document, and then immediately moves to the next row to begin the next merge document.
#         FieldSkipIf fieldSkipIf = (FieldSkipIf) builder.insert_field(FieldType.field_skip_if, true)
#
#         # Move the builder to the SKIPIF field's separator so we can place a MERGEFIELD inside the SKIPIF field.
#         builder.move_to(fieldSkipIf.separator)
#         FieldMergeField fieldMergeField = (FieldMergeField)builder.insert_field(FieldType.field_merge_field, true)
#         fieldMergeField.field_name = "Department"
#
#         # The MERGEFIELD refers to the "Department" column in our data table. If a row from that table
#         # has a value of "HR" in its "Department" column, then this row will fulfill the condition.
#         fieldSkipIf.left_expression = "="
#         fieldSkipIf.right_expression = "HR"
#
#         # Add content to our document, create the data source, and execute the mail merge.
#         builder.move_to_document_end()
#         builder.write("Dear ")
#         fieldMergeField = (FieldMergeField)builder.insert_field(FieldType.field_merge_field, true)
#         fieldMergeField.field_name = "Name"
#         builder.writeln(", ")
#
#         # This table has three rows, and one of them fulfills the condition of our SKIPIF field.
#         # The mail merge will produce two pages.
#         DataTable table = new DataTable("Employees")
#         table.columns.add("Name")
#         table.columns.add("Department")
#         table.rows.add("John Doe", "Sales")
#         table.rows.add("Jane Doe", "Accounting")
#         table.rows.add("John Cardholder", "HR")
#
#         doc.mail_merge.execute(table)
#         doc.save(aeb.artifacts_dir + "Field.skipif.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.skipif.docx")
#
#         self.assertEqual(0, doc.range.fields.count)
#         self.assertEqual("Dear John Doe, \r" +
#                         "\fDear Jane Doe, \r\f", doc.get_text())
#
#
#     def test_field_set_ref(self) :
#
#         #ExStart
#         #ExFor:FieldRef
#         #ExFor:FieldRef.bookmark_name
#         #ExFor:FieldSet
#         #ExFor:FieldSet.bookmark_name
#         #ExFor:FieldSet.bookmark_text
#         #ExSummary:Shows how to create bookmarked text with a SET field, and then display it in the document using a REF field.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Name bookmarked text with a SET field.
#         # This field refers to the "bookmark" not a bookmark structure that appears within the text, but a named variable.
#         FieldSet fieldSet = (FieldSet)builder.insert_field(FieldType.field_set, false)
#         fieldSet.bookmark_name = "MyBookmark"
#         fieldSet.bookmark_text = "Hello world!"
#         fieldSet.update()
#
#         self.assertEqual(" SET  MyBookmark \"Hello world!\"", fieldSet.get_field_code())
#
#         # Refer to the bookmark by name in a REF field and display its contents.
#         FieldRef fieldRef = (FieldRef)builder.insert_field(FieldType.field_ref, true)
#         fieldRef.bookmark_name = "MyBookmark"
#         fieldRef.update()
#
#         self.assertEqual(" REF  MyBookmark", fieldRef.get_field_code())
#         self.assertEqual("Hello world!", fieldRef.result)
#
#         doc.save(aeb.artifacts_dir + "Field.set.ref.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.set.ref.docx")
#
#         self.assertEqual("Hello world!", doc.range.bookmarks[0].text)
#
#         fieldSet = (FieldSet)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_set, " SET  MyBookmark \"Hello world!\"", "Hello world!", fieldSet)
#         self.assertEqual("MyBookmark", fieldSet.bookmark_name)
#         self.assertEqual("Hello world!", fieldSet.bookmark_text)
#
#         TestUtil.verify_field(FieldType.field_ref, " REF  MyBookmark", "Hello world!", fieldRef)
#         self.assertEqual("Hello world!", fieldRef.result)
#
#
#     [Test]
#     [Ignore("WORDSNET-18137")]
#     public void FieldTemplate()
#
#         #ExStart
#         #ExFor:FieldTemplate
#         #ExFor:FieldTemplate.include_full_path
#         #ExSummary:Shows how to use a TEMPLATE field to display the local file system location of a document's template.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         FieldTemplate field = (FieldTemplate)builder.insert_field(FieldType.field_template, false)
#         self.assertEqual(" TEMPLATE ", field.get_field_code())
#
#         builder.writeln()
#         field = (FieldTemplate)builder.insert_field(FieldType.field_template, false)
#         field.include_full_path = true
#
#         self.assertEqual(" TEMPLATE  \\p", field.get_field_code())
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.template.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.template.docx")
#
#         field = (FieldTemplate)doc.range.fields[0]
#         self.assertEqual(" TEMPLATE ", field.get_field_code())
#         self.assertEqual("Normal.dotm", field.result)
#
#         field = (FieldTemplate)doc.range.fields[1]
#         self.assertEqual(" TEMPLATE  \\p", field.get_field_code())
#         self.assertTrue(field.result.ends_with("\\Microsoft\\Templates\\Normal.dotm"))
#
#
#
#     def test_field_symbol(self) :
#
#         #ExStart
#         #ExFor:FieldSymbol
#         #ExFor:FieldSymbol.character_code
#         #ExFor:FieldSymbol.dont_affects_line_spacing
#         #ExFor:FieldSymbol.font_name
#         #ExFor:FieldSymbol.font_size
#         #ExFor:FieldSymbol.is_ansi
#         #ExFor:FieldSymbol.is_shift_jis
#         #ExFor:FieldSymbol.is_unicode
#         #ExSummary:Shows how to use the SYMBOL field.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Below are three ways to use a SYMBOL field to display a single character.
#         # 1 -  Add a SYMBOL field which displays the © (Copyright) symbol, specified by an ANSI character code:
#         FieldSymbol field = (FieldSymbol)builder.insert_field(FieldType.field_symbol, true)
#
#         # The ANSI character code "U+00A9", or "169" in integer form, is reserved for the copyright symbol.
#         field.character_code = 0x00a9.to_string()
#         field.is_ansi = true
#
#         self.assertEqual(" SYMBOL  169 \\a", field.get_field_code())
#
#         builder.writeln(" Line 1")
#
#         # 2 -  Add a SYMBOL field which displays the ∞ (Infinity) symbol, and modify its appearance:
#         field = (FieldSymbol)builder.insert_field(FieldType.field_symbol, true)
#
#         # In Unicode, the infinity symbol occupies the "221E" code.
#         field.character_code = 0x221E.to_string()
#         field.is_unicode = true
#
#         # Change the font of our symbol after using the Windows Character Map
#         # to ensure that the font can represent that symbol.
#         field.font_name = "Calibri"
#         field.font_size = "24"
#
#         # We can set this flag for tall symbols to make them not push down the rest of the text on their line.
#         field.dont_affects_line_spacing = true
#
#         self.assertEqual(" SYMBOL  8734 \\u \\f Calibri \\s 24 \\h", field.get_field_code())
#
#         builder.writeln("Line 2")
#
#         # 3 -  Add a SYMBOL field which displays the あ character,
#         # with a font that supports Shift-JIS (Windows-932) codepage:
#         field = (FieldSymbol)builder.insert_field(FieldType.field_symbol, true)
#         field.font_name = "MS Gothic"
#         field.character_code = 0x82A0.to_string()
#         field.is_shift_jis = true
#
#         self.assertEqual(" SYMBOL  33440 \\f \"MS Gothic\" \\j", field.get_field_code())
#
#         builder.write("Line 3")
#
#         doc.save(aeb.artifacts_dir + "Field.symbol.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.symbol.docx")
#
#         field = (FieldSymbol)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_symbol, " SYMBOL  169 \\a", string.empty, field)
#         self.assertEqual(0x00a9.to_string(), field.character_code)
#         self.assertTrue(field.is_ansi)
#         self.assertEqual("©", field.display_result)
#
#         field = (FieldSymbol)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_symbol, " SYMBOL  8734 \\u \\f Calibri \\s 24 \\h", string.empty, field)
#         self.assertEqual(0x221E.to_string(), field.character_code)
#         self.assertEqual("Calibri", field.font_name)
#         self.assertEqual("24", field.font_size)
#         self.assertTrue(field.is_unicode)
#         self.assertTrue(field.dont_affects_line_spacing)
#         self.assertEqual("∞", field.display_result)
#
#         field = (FieldSymbol)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_symbol, " SYMBOL  33440 \\f \"MS Gothic\" \\j", string.empty, field)
#         self.assertEqual(0x82A0.to_string(), field.character_code)
#         self.assertEqual("MS Gothic", field.font_name)
#         self.assertTrue(field.is_shift_jis)
#
#
#     def test_field_title(self) :
#
#         #ExStart
#         #ExFor:FieldTitle
#         #ExFor:FieldTitle.text
#         #ExSummary:Shows how to use the TITLE field.
#         doc = aw.Document()
#
#         # Set a value for the "Title" built-in document property.
#         doc.built_in_document_properties.title = "My Title"
#
#         # We can use the TITLE field to display the value of this property in the document.
#         builder = aw.DocumentBuilder(doc)
#         FieldTitle field = (FieldTitle)builder.insert_field(FieldType.field_title, false)
#         field.update()
#
#         self.assertEqual(" TITLE ", field.get_field_code())
#         self.assertEqual("My Title", field.result)
#
#         # Setting a value for the field's Text property,
#         # and then updating the field will also overwrite the corresponding built-in property with the new value.
#         builder.writeln()
#         field = (FieldTitle)builder.insert_field(FieldType.field_title, false)
#         field.text = "My New Title"
#         field.update()
#
#         self.assertEqual(" TITLE  \"My New Title\"", field.get_field_code())
#         self.assertEqual("My New Title", field.result)
#         self.assertEqual("My New Title", doc.built_in_document_properties.title)
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.title.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.title.docx")
#
#         self.assertEqual("My New Title", doc.built_in_document_properties.title)
#
#         field = (FieldTitle)doc.range.fields[0]
#
#         TestUtil.verify_field(FieldType.field_title, " TITLE ", "My New Title", field)
#
#         field = (FieldTitle)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_title, " TITLE  \"My New Title\"", "My New Title", field)
#         self.assertEqual("My New Title", field.text)
#
#
#     #ExStart
#     #ExFor:FieldToa
#     #ExFor:FieldToa.bookmark_name
#     #ExFor:FieldToa.entry_category
#     #ExFor:FieldToa.entry_separator
#     #ExFor:FieldToa.page_number_list_separator
#     #ExFor:FieldToa.page_range_separator
#     #ExFor:FieldToa.remove_entry_formatting
#     #ExFor:FieldToa.sequence_name
#     #ExFor:FieldToa.sequence_separator
#     #ExFor:FieldToa.use_heading
#     #ExFor:FieldToa.use_passim
#     #ExFor:FieldTA
#     #ExFor:FieldTA.entry_category
#     #ExFor:FieldTA.is_bold
#     #ExFor:FieldTA.is_italic
#     #ExFor:FieldTA.long_citation
#     #ExFor:FieldTA.page_range_bookmark_name
#     #ExFor:FieldTA.short_citation
#     #ExSummary:Shows how to build and customize a table of authorities using TOA and TA fields.
#     def test_field_toa(self) :
#
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # Insert a TOA field, which will create an entry for each TA field in the document,
#         # displaying long citations and page numbers for each entry.
#         FieldToa fieldToa = (FieldToa)builder.insert_field(FieldType.field_toa, false)
#
#         # Set the entry category for our table. This TOA will now only include TA fields
#         # that have a matching value in their EntryCategory property.
#         fieldToa.entry_category = "1"
#
#         # Moreover, the Table of Authorities category at index 1 is "Cases",
#         # which will show up as our table's title if we set this variable to true.
#         fieldToa.use_heading = true
#
#         # We can further filter TA fields by naming a bookmark that they will need to be within the TOA bounds.
#         fieldToa.bookmark_name = "MyBookmark"
#
#         # By default, a dotted line page-wide tab appears between the TA field's citation
#         # and its page number. We can replace it with any text we put on this property.
#         # Inserting a tab character will preserve the original tab.
#         fieldToa.entry_separator = " \t p."
#
#         # If we have multiple TA entries that share the same long citation,
#         # all their respective page numbers will show up on one row.
#         # We can use this property to specify a string that will separate their page numbers.
#         fieldToa.page_number_list_separator = " & p. "
#
#         # We can set this to true to get our table to display the word "passim"
#         # if there are five or more page numbers in one row.
#         fieldToa.use_passim = true
#
#         # One TA field can refer to a range of pages.
#         # We can specify a string here to appear between the start and end page numbers for such ranges.
#         fieldToa.page_range_separator = " to "
#
#         # The format from the TA fields will carry over into our table.
#         # We can disable this by setting the RemoveEntryFormatting flag.
#         fieldToa.remove_entry_formatting = true
#         builder.font.color = Color.green
#         builder.font.name = "Arial Black"
#
#         self.assertEqual(" TOA  \\c 1 \\h \\b MyBookmark \\e \" \t p.\" \\l \" & p. \" \\p \\g \" to \" \\f", fieldToa.get_field_code())
#
#         builder.insert_break(BreakType.page_break)
#
#         # This TA field will not appear as an entry in the TOA since it is outside
#         # the bookmark's bounds that the TOA's BookmarkName property specifies.
#         FieldTA fieldTA = InsertToaEntry(builder, "1", "Source 1")
#
#         self.assertEqual(" TA  \\c 1 \\l \"Source 1\"", fieldTA.get_field_code())
#
#         # This TA field is inside the bookmark,
#         # but the entry category does not match that of the table, so the TA field will not include it.
#         builder.start_bookmark("MyBookmark")
#         fieldTA = InsertToaEntry(builder, "2", "Source 2")
#
#         # This entry will appear in the table.
#         fieldTA = InsertToaEntry(builder, "1", "Source 3")
#
#         # A TOA table does not display short citations,
#         # but we can use them as a shorthand to refer to bulky source names that multiple TA fields reference.
#         fieldTA.short_citation = "S.3"
#
#         self.assertEqual(" TA  \\c 1 \\l \"Source 3\" \\s S.3", fieldTA.get_field_code())
#
#         # We can format the page number to make it bold/italic using the following properties.
#         # We will still see these effects if we set our table to ignore formatting.
#         fieldTA = InsertToaEntry(builder, "1", "Source 2")
#         fieldTA.is_bold = true
#         fieldTA.is_italic = true
#
#         self.assertEqual(" TA  \\c 1 \\l \"Source 2\" \\b \\i", fieldTA.get_field_code())
#
#         # We can configure TA fields to get their TOA entries to refer to a range of pages that a bookmark spans across.
#         # Note that this entry refers to the same source as the one above to share one row in our table.
#         # This row will have the page number of the entry above and the page range of this entry,
#         # with the table's page list and page number range separators between page numbers.
#         fieldTA = InsertToaEntry(builder, "1", "Source 3")
#         fieldTA.page_range_bookmark_name = "MyMultiPageBookmark"
#
#         builder.start_bookmark("MyMultiPageBookmark")
#         builder.insert_break(BreakType.page_break)
#         builder.insert_break(BreakType.page_break)
#         builder.insert_break(BreakType.page_break)
#         builder.end_bookmark("MyMultiPageBookmark")
#
#         self.assertEqual(" TA  \\c 1 \\l \"Source 3\" \\r MyMultiPageBookmark", fieldTA.get_field_code())
#
#         # If we have enabled the "Passim" feature of our table, having 5 or more TA entries with the same source will invoke it.
#         for (int i = 0 i < 5 i++)
#
#             InsertToaEntry(builder, "1", "Source 4")
#
#
#         builder.end_bookmark("MyBookmark")
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.toa.ta.docx")
#         TestFieldTOA(new Document(aeb.artifacts_dir + "Field.toa.ta.docx")) #ExSKip
#
#
#     private static FieldTA InsertToaEntry(DocumentBuilder builder, string entryCategory, string longCitation)
#
#         FieldTA field = (FieldTA)builder.insert_field(FieldType.field_toa_entry, false)
#         field.entry_category = entryCategory
#         field.long_citation = longCitation
#
#         builder.insert_break(BreakType.page_break)
#
#         return field
#
#     #ExEnd
#
#     private void TestFieldTOA(Document doc)
#
#         FieldToa fieldTOA = (FieldToa)doc.range.fields[0]
#
#         self.assertEqual("1", fieldTOA.entry_category)
#         self.assertTrue(fieldTOA.use_heading)
#         self.assertEqual("MyBookmark", fieldTOA.bookmark_name)
#         self.assertEqual(" \t p.", fieldTOA.entry_separator)
#         self.assertEqual(" & p. ", fieldTOA.page_number_list_separator)
#         self.assertTrue(fieldTOA.use_passim)
#         self.assertEqual(" to ", fieldTOA.page_range_separator)
#         self.assertTrue(fieldTOA.remove_entry_formatting)
#         self.assertEqual(" TOA  \\c 1 \\h \\b MyBookmark \\e \" \t p.\" \\l \" & p. \" \\p \\g \" to \" \\f", fieldTOA.get_field_code())
#         self.assertEqual("Cases\r" +
#                         "Source 2 \t p.5\r" +
#                         "Source 3 \t p.4 & p. 7 to 10\r" +
#                         "Source 4 \t p.passim\r", fieldTOA.result)
#
#         FieldTA fieldTA = (FieldTA)doc.range.fields[1]
#
#         TestUtil.verify_field(FieldType.field_toa_entry, " TA  \\c 1 \\l \"Source 1\"", string.empty, fieldTA)
#         self.assertEqual("1", fieldTA.entry_category)
#         self.assertEqual("Source 1", fieldTA.long_citation)
#
#         fieldTA = (FieldTA)doc.range.fields[2]
#
#         TestUtil.verify_field(FieldType.field_toa_entry, " TA  \\c 2 \\l \"Source 2\"", string.empty, fieldTA)
#         self.assertEqual("2", fieldTA.entry_category)
#         self.assertEqual("Source 2", fieldTA.long_citation)
#
#         fieldTA = (FieldTA)doc.range.fields[3]
#
#         TestUtil.verify_field(FieldType.field_toa_entry, " TA  \\c 1 \\l \"Source 3\" \\s S.3", string.empty, fieldTA)
#         self.assertEqual("1", fieldTA.entry_category)
#         self.assertEqual("Source 3", fieldTA.long_citation)
#         self.assertEqual("S.3", fieldTA.short_citation)
#
#         fieldTA = (FieldTA)doc.range.fields[4]
#
#         TestUtil.verify_field(FieldType.field_toa_entry, " TA  \\c 1 \\l \"Source 2\" \\b \\i", string.empty, fieldTA)
#         self.assertEqual("1", fieldTA.entry_category)
#         self.assertEqual("Source 2", fieldTA.long_citation)
#         self.assertTrue(fieldTA.is_bold)
#         self.assertTrue(fieldTA.is_italic)
#
#         fieldTA = (FieldTA)doc.range.fields[5]
#
#         TestUtil.verify_field(FieldType.field_toa_entry, " TA  \\c 1 \\l \"Source 3\" \\r MyMultiPageBookmark", string.empty, fieldTA)
#         self.assertEqual("1", fieldTA.entry_category)
#         self.assertEqual("Source 3", fieldTA.long_citation)
#         self.assertEqual("MyMultiPageBookmark", fieldTA.page_range_bookmark_name)
#
#         for (int i = 6 i < 11 i++)
#
#             fieldTA = (FieldTA)doc.range.fields[i]
#
#             TestUtil.verify_field(FieldType.field_toa_entry, " TA  \\c 1 \\l \"Source 4\"", string.empty, fieldTA)
#             self.assertEqual("1", fieldTA.entry_category)
#             self.assertEqual("Source 4", fieldTA.long_citation)
#
#
#
#     def test_field_add_in(self) :
#
#         #ExStart
#         #ExFor:FieldAddIn
#         #ExSummary:Shows how to process an ADDIN field.
#         Document doc = new Document(aeb.my_dir + "Field sample - ADDIN.docx")
#
#         # Aspose.words does not support inserting ADDIN fields, but we can still load and read them.
#         FieldAddIn field = (FieldAddIn)doc.range.fields[0]
#
#         self.assertEqual(" ADDIN \"My value\" ", field.get_field_code())
#         #ExEnd
#
#         doc = DocumentHelper.save_open(doc)
#
#         TestUtil.verify_field(FieldType.field_addin, " ADDIN \"My value\" ", string.empty, doc.range.fields[0])
#
#
#     def test_field_edit_time(self) :
#
#         #ExStart
#         #ExFor:FieldEditTime
#         #ExSummary:Shows how to use the EDITTIME field.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # The EDITTIME field will show, in minutes,
#         # the time spent with the document open in a Microsoft Word window.
#         builder.move_to_header_footer(HeaderFooterType.header_primary)
#         builder.write("You've been editing this document for ")
#         FieldEditTime field = (FieldEditTime)builder.insert_field(FieldType.field_edit_time, true)
#         builder.writeln(" minutes.")
#
#         # This built in document property tracks the minutes. Microsoft Word uses this property
#         # to track the time spent with the document open. We can also edit it ourselves.
#         doc.built_in_document_properties.total_editing_time = 10
#         field.update()
#
#         self.assertEqual(" EDITTIME ", field.get_field_code())
#         self.assertEqual("10", field.result)
#
#         # The field does not update itself in real-time, and will also have to be
#         # manually updated in Microsoft Word anytime we need an accurate value.
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.edittime.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.edittime.docx")
#
#         self.assertEqual(10, doc.built_in_document_properties.total_editing_time)
#
#         TestUtil.verify_field(FieldType.field_edit_time, " EDITTIME ", "10", doc.range.fields[0])
#
#
#     #ExStart
#     #ExFor:FieldEQ
#     #ExSummary:Shows how to use the EQ field to display a variety of mathematical equations.
#     def test_field_eq(self) :
#
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # An EQ field displays a mathematical equation consisting of one or many elements.
#         # Each element takes the following form: [switch][options][arguments].
#         # There may be one switch, and several possible options.
#         # The arguments are a set of coma-separated values enclosed by round braces.
#
#         # Here we use a document builder to insert an EQ field, with an "\f" switch, which corresponds to "Fraction".
#         # We will pass values 1 and 4 as arguments, and we will not use any options.
#         # This field will display a fraction with 1 as the numerator and 4 as the denominator.
#         FieldEQ field = InsertFieldEQ(builder, @"\f(1,4)")
#
#         self.assertEqual(@" EQ \f(1,4)", field.get_field_code())
#
#         # One EQ field may contain multiple elements placed sequentially.
#         # We can also nest elements inside one another by placing the inner elements
#         # inside the argument brackets of outer elements.
#         # We can find the full list of switches, along with their uses here:
#         # https:#blogs.msdn.microsoft.com/murrays/2018/01/23/microsoft-word-eq-field/
#
#         # Below are applications of nine different EQ field switches that we can use to create different kinds of objects.
#         # 1 -  Array switch "\a", aligned left, 2 columns, 3 points of horizontal and vertical spacing:
#         InsertFieldEQ(builder, @"\a \al \co2 \vs3 \hs3(4x,- 4y,-4x,+ y)")
#
#         # 2 -  Bracket switch "\b", bracket character "[", to enclose the contents in a set of square braces:
#         # Note that we are nesting an array inside the brackets, which will altogether look like a matrix in the output.
#         InsertFieldEQ(builder, @"\b \bc\[ (\a \al \co3 \vs3 \hs3(1,0,0,0,1,0,0,0,1))")
#
#         # 3 -  Displacement switch "\d", displacing text "B" 30 spaces to the right of "A", displaying the gap as an underline:
#         InsertFieldEQ(builder, @"A \d \fo30 \li() B")
#
#         # 4 -  Formula consisting of multiple fractions:
#         InsertFieldEQ(builder, @"\f(d,dx)(u + v) = \f(du,dx) + \f(dv,dx)")
#
#         # 5 -  Integral switch "\i", with a summation symbol:
#         InsertFieldEQ(builder, @"\i \su(n=1,5,n)")
#
#         # 6 -  List switch "\l":
#         InsertFieldEQ(builder, @"\l(1,1,2,3,n,8,13)")
#
#         # 7 -  Radical switch "\r", displaying a cubed root of x:
#         InsertFieldEQ(builder, @"\r (3,x)")
#
#         # 8 -  Subscript/superscript switch "/s", first as a superscript and then as a subscript:
#         InsertFieldEQ(builder, @"\s \up8(Superscript) Text \s \do8(Subscript)")
#
#         # 9 -  Box switch "\x", with lines at the top, bottom, left and right of the input:
#         InsertFieldEQ(builder, @"\x \to \bo \le \ri(5)")
#
#         # Some more complex combinations.
#         InsertFieldEQ(builder, @"\a \ac \vs1 \co1(lim,n→∞) \b (\f(n,n2 + 12) + \f(n,n2 + 22) + ... + \f(n,n2 + n2))")
#         InsertFieldEQ(builder, @"\i (,,  \b(\f(x,x2 + 3x + 2))) \s \up10(2)")
#         InsertFieldEQ(builder, @"\i \in( tan x, \s \up2(sec x), \b(\r(3) )\s \up4(t) \s \up7(2)  dt)")
#
#         doc.save(aeb.artifacts_dir + "Field.eq.docx")
#         TestFieldEQ(new Document(aeb.artifacts_dir + "Field.eq.docx")) #ExSkip
#
#
#     # <summary>
#     # Use a document builder to insert an EQ field, set its arguments and start a new paragraph.
#     # </summary>
#     private static FieldEQ InsertFieldEQ(DocumentBuilder builder, string args)
#
#         FieldEQ field = (FieldEQ)builder.insert_field(FieldType.field_equation, true)
#         builder.move_to(field.separator)
#         builder.write(args)
#         builder.move_to(field.start.parent_node)
#
#         builder.insert_paragraph()
#         return field
#
#     #ExEnd
#
#     private void TestFieldEQ(Document doc)
#
#         TestUtil.verify_field(FieldType.field_equation, @" EQ \f(1,4)", string.empty, doc.range.fields[0])
#         TestUtil.verify_field(FieldType.field_equation, @" EQ \a \al \co2 \vs3 \hs3(4x,- 4y,-4x,+ y)", string.empty, doc.range.fields[1])
#         TestUtil.verify_field(FieldType.field_equation, @" EQ \b \bc\[ (\a \al \co3 \vs3 \hs3(1,0,0,0,1,0,0,0,1))", string.empty, doc.range.fields[2])
#         TestUtil.verify_field(FieldType.field_equation, @" EQ A \d \fo30 \li() B", string.empty, doc.range.fields[3])
#         TestUtil.verify_field(FieldType.field_equation, @" EQ \f(d,dx)(u + v) = \f(du,dx) + \f(dv,dx)", string.empty, doc.range.fields[4])
#         TestUtil.verify_field(FieldType.field_equation, @" EQ \i \su(n=1,5,n)", string.empty, doc.range.fields[5])
#         TestUtil.verify_field(FieldType.field_equation, @" EQ \l(1,1,2,3,n,8,13)", string.empty, doc.range.fields[6])
#         TestUtil.verify_field(FieldType.field_equation, @" EQ \r (3,x)", string.empty, doc.range.fields[7])
#         TestUtil.verify_field(FieldType.field_equation, @" EQ \s \up8(Superscript) Text \s \do8(Subscript)", string.empty, doc.range.fields[8])
#         TestUtil.verify_field(FieldType.field_equation, @" EQ \x \to \bo \le \ri(5)", string.empty, doc.range.fields[9])
#         TestUtil.verify_field(FieldType.field_equation, @" EQ \a \ac \vs1 \co1(lim,n→∞) \b (\f(n,n2 + 12) + \f(n,n2 + 22) + ... + \f(n,n2 + n2))", string.empty, doc.range.fields[10])
#         TestUtil.verify_field(FieldType.field_equation, @" EQ \i (,,  \b(\f(x,x2 + 3x + 2))) \s \up10(2)", string.empty, doc.range.fields[11])
#         TestUtil.verify_field(FieldType.field_equation, @" EQ \i \in( tan x, \s \up2(sec x), \b(\r(3) )\s \up4(t) \s \up7(2)  dt)", string.empty, doc.range.fields[12])
#         TestUtil.verify_web_response_status_code(HttpStatusCode.ok, "https:#blogs.msdn.microsoft.com/murrays/2018/01/23/microsoft-word-eq-field/")
#
#
#     def test_field_forms(self) :
#
#         #ExStart
#         #ExFor:FieldFormCheckBox
#         #ExFor:FieldFormDropDown
#         #ExFor:FieldFormText
#         #ExSummary:Shows how to process FORMCHECKBOX, FORMDROPDOWN and FORMTEXT fields.
#         # These fields are legacy equivalents of the FormField. We can read, but not create these fields using Aspose.words.
#         # In Microsoft Word, we can insert these fields via the Legacy Tools menu in the Developer tab.
#         Document doc = new Document(aeb.my_dir + "Form fields.docx")
#
#         FieldFormCheckBox fieldFormCheckBox = (FieldFormCheckBox)doc.range.fields[1]
#         self.assertEqual(" FORMCHECKBOX \u0001", fieldFormCheckBox.get_field_code())
#
#         FieldFormDropDown fieldFormDropDown = (FieldFormDropDown)doc.range.fields[2]
#         self.assertEqual(" FORMDROPDOWN \u0001", fieldFormDropDown.get_field_code())
#
#         FieldFormText fieldFormText = (FieldFormText)doc.range.fields[0]
#         self.assertEqual(" FORMTEXT \u0001", fieldFormText.get_field_code())
#         #ExEnd
#
#
#     def test_field_formula(self) :
#
#         #ExStart
#         #ExFor:FieldFormula
#         #ExSummary:Shows how to use the formula field to display the result of an equation.
#         doc = aw.Document()
#
#         # Use a field builder to construct a mathematical equation,
#         # then create a formula field to display the equation's result in the document.
#         FieldBuilder fieldBuilder = new FieldBuilder(FieldType.field_formula)
#         fieldBuilder.add_argument(2)
#         fieldBuilder.add_argument("*")
#         fieldBuilder.add_argument(5)
#
#         FieldFormula field = (FieldFormula)fieldBuilder.build_and_insert(doc.first_section.body.first_paragraph)
#         field.update()
#
#         self.assertEqual(" = 2 * 5 ", field.get_field_code())
#         self.assertEqual("10", field.result)
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.formula.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.formula.docx")
#
#         TestUtil.verify_field(FieldType.field_formula, " = 2 * 5 ", "10", doc.range.fields[0])
#
#
#     def test_field_last_saved_by(self) :
#
#         #ExStart
#         #ExFor:FieldLastSavedBy
#         #ExSummary:Shows how to use the LASTSAVEDBY field.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # If we create a document in Microsoft Word, it will have the user's name in the "Last saved by" built-in property.
#         # If we make a document programmatically, this property will be null, and we will need to assign a value.
#         doc.built_in_document_properties.last_saved_by = "John Doe"
#
#         # We can use the LASTSAVEDBY field to display the value of this property in the document.
#         FieldLastSavedBy field = (FieldLastSavedBy)builder.insert_field(FieldType.field_last_saved_by, true)
#
#         self.assertEqual(" LASTSAVEDBY ", field.get_field_code())
#         self.assertEqual("John Doe", field.result)
#
#         doc.save(aeb.artifacts_dir + "Field.lastsavedby.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.lastsavedby.docx")
#
#         self.assertEqual("John Doe", doc.built_in_document_properties.last_saved_by)
#         TestUtil.verify_field(FieldType.field_last_saved_by, " LASTSAVEDBY ", "John Doe", doc.range.fields[0])
#
#
#     [Test]
#     [Ignore("WORDSNET-18173")]
#     public void FieldMergeRec()
#
#         #ExStart
#         #ExFor:FieldMergeRec
#         #ExFor:FieldMergeSeq
#         #ExFor:FieldSkipIf
#         #ExFor:FieldSkipIf.comparison_operator
#         #ExFor:FieldSkipIf.left_expression
#         #ExFor:FieldSkipIf.right_expression
#         #ExSummary:Shows how to use MERGEREC and MERGESEQ fields to the number and count mail merge records in a mail merge's output documents.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         builder.write("Dear ")
#         FieldMergeField fieldMergeField = (FieldMergeField)builder.insert_field(FieldType.field_merge_field, true)
#         fieldMergeField.field_name = "Name"
#         builder.writeln(",")
#
#         # A MERGEREC field will print the row number of the data being merged in every merge output document.
#         builder.write("\nRow number of record in data source: ")
#         FieldMergeRec fieldMergeRec = (FieldMergeRec)builder.insert_field(FieldType.field_merge_rec, true)
#
#         self.assertEqual(" MERGEREC ", fieldMergeRec.get_field_code())
#
#         # A MERGESEQ field will count the number of successful merges and print the current value on each respective page.
#         # If a mail merge skips no rows and invokes no SKIP/SKIPIF/NEXT/NEXTIF fields, then all merges are successful.
#         # The MERGESEQ and MERGEREC fields will display the same results of their mail merge was successful.
#         builder.write("\nSuccessful merge number: ")
#         FieldMergeSeq fieldMergeSeq = (FieldMergeSeq)builder.insert_field(FieldType.field_merge_seq, true)
#
#         self.assertEqual(" MERGESEQ ", fieldMergeSeq.get_field_code())
#
#         # Insert a SKIPIF field, which will skip a merge if the name is "John Doe".
#         FieldSkipIf fieldSkipIf = (FieldSkipIf)builder.insert_field(FieldType.field_skip_if, true)
#         builder.move_to(fieldSkipIf.separator)
#         fieldMergeField = (FieldMergeField)builder.insert_field(FieldType.field_merge_field, true)
#         fieldMergeField.field_name = "Name"
#         fieldSkipIf.left_expression = "="
#         fieldSkipIf.right_expression = "John Doe"
#
#         # Create a data source with 3 rows, one of them having "John Doe" as a value for the "Name" column.
#         # Since a SKIPIF field will be triggered once by that value, the output of our mail merge will have 2 pages instead of 3.
#         # On page 1, the MERGESEQ and MERGEREC fields will both display "1".
#         # On page 2, the MERGEREC field will display "3" and the MERGESEQ field will display "2".
#         DataTable table = new DataTable("Employees")
#         table.columns.add("Name")
#         table.rows.add(new[]  "Jane Doe" )
#         table.rows.add(new[]  "John Doe" )
#         table.rows.add(new[]  "Joe Bloggs" )
#
#         doc.mail_merge.execute(table)
#         doc.save(aeb.artifacts_dir + "Field.mergerec.mergeseq.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.mergerec.mergeseq.docx")
#
#         self.assertEqual(0, doc.range.fields.count)
#
#         self.assertEqual("Dear Jane Doe,\r" +
#                         "\r" +
#                         "Row number of record in data source: 1\r" +
#                         "Successful merge number: 1\fDear Joe Bloggs,\r" +
#                         "\r" +
#                         "Row number of record in data source: 2\r" +
#                         "Successful merge number: 3", doc.get_text().strip())
#
#
#     def test_field_ocx(self) :
#
#         #ExStart
#         #ExFor:FieldOcx
#         #ExSummary:Shows how to insert an OCX field.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         FieldOcx field = (FieldOcx)builder.insert_field(FieldType.field_ocx, true)
#
#         self.assertEqual(" OCX ", field.get_field_code())
#         #ExEnd
#
#         TestUtil.verify_field(FieldType.field_ocx, " OCX ", string.empty, field)
#
#
#     #ExStart
#     #ExFor:Field.remove
#     #ExFor:FieldPrivate
#     #ExSummary:Shows how to process PRIVATE fields.
#     def test_field_private(self) :
#
#         # Open a Corel WordPerfect document which we have converted to .docx format.
#         Document doc = new Document(aeb.my_dir + "Field sample - PRIVATE.docx")
#
#         # WordPerfect 5.x/6.x documents like the one we have loaded may contain PRIVATE fields.
#         # Microsoft Word preserves PRIVATE fields during load/save operations,
#         # but provides no functionality for them.
#         FieldPrivate field = (FieldPrivate)doc.range.fields[0]
#
#         self.assertEqual(" PRIVATE \"My value\" ", field.get_field_code())
#         self.assertEqual(FieldType.field_private, field.type)
#
#         # We can also insert PRIVATE fields using a document builder.
#         builder = aw.DocumentBuilder(doc)
#         builder.insert_field(FieldType.field_private, true)
#
#         # These fields are not a viable way of protecting sensitive information.
#         # Unless backward compatibility with older versions of WordPerfect is essential,
#         # we can safely remove these fields. We can do this using a DocumentVisiitor implementation.
#         self.assertEqual(2, doc.range.fields.count)
#
#         FieldPrivateRemover remover = new FieldPrivateRemover()
#         doc.accept(remover)
#
#         self.assertEqual(2, remover.get_fields_removed_count())
#         self.assertEqual(0, doc.range.fields.count)
#
#
#     # <summary>
#     # Removes all encountered PRIVATE fields.
#     # </summary>
#     public class FieldPrivateRemover : DocumentVisitor
#
#         public FieldPrivateRemover()
#
#             mFieldsRemovedCount = 0
#
#
#         public int GetFieldsRemovedCount()
#
#             return mFieldsRemovedCount
#
#
#         # <summary>
#         # Called when a FieldEnd node is encountered in the document.
#         # If the node belongs to a PRIVATE field, the entire field is removed.
#         # </summary>
#         public override VisitorAction VisitFieldEnd(FieldEnd fieldEnd)
#
#             if (fieldEnd.field_type == FieldType.field_private)
#
#                 fieldEnd.get_field().remove()
#                 mFieldsRemovedCount++
#
#
#             return VisitorAction.continue
#
#
#         private int mFieldsRemovedCount
#
#     #ExEnd
#
#     def test_field_section(self) :
#
#         #ExStart
#         #ExFor:FieldSection
#         #ExFor:FieldSectionPages
#         #ExSummary:Shows how to use SECTION and SECTIONPAGES fields to number pages by sections.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         builder.move_to_header_footer(HeaderFooterType.header_primary)
#         builder.paragraph_format.alignment = ParagraphAlignment.right
#
#         # A SECTION field displays the number of the section it is in.
#         builder.write("Section ")
#         FieldSection fieldSection = (FieldSection)builder.insert_field(FieldType.field_section, true)
#
#         self.assertEqual(" SECTION ", fieldSection.get_field_code())
#
#         # A PAGE field displays the number of the page it is in.
#         builder.write("\nPage ")
#         FieldPage fieldPage = (FieldPage)builder.insert_field(FieldType.field_page, true)
#
#         self.assertEqual(" PAGE ", fieldPage.get_field_code())
#
#         # A SECTIONPAGES field displays the number of pages that the section it is in spans across.
#         builder.write(" of ")
#         FieldSectionPages fieldSectionPages = (FieldSectionPages)builder.insert_field(FieldType.field_section_pages, true)
#
#         self.assertEqual(" SECTIONPAGES ", fieldSectionPages.get_field_code())
#
#         # Move out of the header back into the main document and insert two pages.
#         # All these pages will be in the first section. Our fields, which appear once every header,
#         # will number the current/total pages of this section.
#         builder.move_to_document_end()
#         builder.insert_break(BreakType.page_break)
#         builder.insert_break(BreakType.page_break)
#
#         # We can insert a new section with the document builder like this.
#         # This will affect the values displayed in the SECTION and SECTIONPAGES fields in all upcoming headers.
#         builder.insert_break(BreakType.section_break_new_page)
#
#         # The PAGE field will keep counting pages across the whole document.
#         # We can manually reset its count at each section to keep track of pages section-by-section.
#         builder.current_section.page_setup.restart_page_numbering = true
#         builder.insert_break(BreakType.page_break)
#
#         doc.update_fields()
#         doc.save(aeb.artifacts_dir + "Field.section.sectionpages.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.section.sectionpages.docx")
#
#         TestUtil.verify_field(FieldType.field_section, " SECTION ", "2", doc.range.fields[0])
#         TestUtil.verify_field(FieldType.field_page, " PAGE ", "2", doc.range.fields[1])
#         TestUtil.verify_field(FieldType.field_section_pages, " SECTIONPAGES ", "2", doc.range.fields[2])
#
#
#     #ExStart
#     #ExFor:FieldTime
#     #ExSummary:Shows how to display the current time using the TIME field.
#     def test_field_time(self) :
#
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # By default, time is displayed in the "h:mm am/pm" format.
#         FieldTime field = InsertFieldTime(builder, "")
#
#         self.assertEqual(" TIME ", field.get_field_code())
#
#         # We can use the \@ flag to change the format of our displayed time.
#         field = InsertFieldTime(builder, "\\@ HHmm")
#
#         self.assertEqual(" TIME \\@ HHmm", field.get_field_code())
#
#         # We can adjust the format to get TIME field to also display the date, according to the Gregorian calendar.
#         field = InsertFieldTime(builder, "\\@ \"M/d/yyyy h mm:ss am/pm\"")
#
#         self.assertEqual(" TIME \\@ \"M/d/yyyy h mm:ss am/pm\"", field.get_field_code())
#
#         doc.save(aeb.artifacts_dir + "Field.time.docx")
#         TestFieldTime(new Document(aeb.artifacts_dir + "Field.time.docx")) #ExSkip
#
#
#     # <summary>
#     # Use a document builder to insert a TIME field, insert a new paragraph and return the field.
#     # </summary>
#     private static FieldTime InsertFieldTime(DocumentBuilder builder, string format)
#
#         FieldTime field = (FieldTime)builder.insert_field(FieldType.field_time, true)
#         builder.move_to(field.separator)
#         builder.write(format)
#         builder.move_to(field.start.parent_node)
#
#         builder.insert_paragraph()
#         return field
#
#     #ExEnd
#
#     private void TestFieldTime(Document doc)
#
#         DateTime docLoadingTime = DateTime.now
#         doc = DocumentHelper.save_open(doc)
#
#         FieldTime field = (FieldTime)doc.range.fields[0]
#
#         self.assertEqual(" TIME ", field.get_field_code())
#         self.assertEqual(FieldType.field_time, field.type)
#         self.assertEqual(DateTime.parse(field.result), DateTime.today.add_hours(docLoadingTime.hour).add_minutes(docLoadingTime.minute))
#
#         field = (FieldTime)doc.range.fields[1]
#
#         self.assertEqual(" TIME \\@ HHmm", field.get_field_code())
#         self.assertEqual(FieldType.field_time, field.type)
#         self.assertEqual(DateTime.parse(field.result), DateTime.today.add_hours(docLoadingTime.hour).add_minutes(docLoadingTime.minute))
#
#         field = (FieldTime)doc.range.fields[2]
#
#         self.assertEqual(" TIME \\@ \"M/d/yyyy h mm:ss am/pm\"", field.get_field_code())
#         self.assertEqual(FieldType.field_time, field.type)
#         self.assertEqual(DateTime.parse(field.result), DateTime.today.add_hours(docLoadingTime.hour).add_minutes(docLoadingTime.minute))
#
#
#     def test_bidi_outline(self) :
#
#         #ExStart
#         #ExFor:FieldBidiOutline
#         #ExFor:FieldShape
#         #ExFor:FieldShape.text
#         #ExFor:ParagraphFormat.bidi
#         #ExSummary:Shows how to create right-to-left language-compatible lists with BIDIOUTLINE fields.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#
#         # The BIDIOUTLINE field numbers paragraphs like the AUTONUM/LISTNUM fields,
#         # but is only visible when a right-to-left editing language is enabled, such as Hebrew or Arabic.
#         # The following field will display ".1", the RTL equivalent of list number "1.".
#         FieldBidiOutline field = (FieldBidiOutline)builder.insert_field(FieldType.field_bidi_outline, true)
#         builder.writeln("שלום")
#
#         self.assertEqual(" BIDIOUTLINE ", field.get_field_code())
#
#         # Add two more BIDIOUTLINE fields, which will display ".2" and ".3".
#         builder.insert_field(FieldType.field_bidi_outline, true)
#         builder.writeln("שלום")
#         builder.insert_field(FieldType.field_bidi_outline, true)
#         builder.writeln("שלום")
#
#         # Set the horizontal text alignment for every paragraph in the document to RTL.
#         foreach (Paragraph para in doc.get_child_nodes(NodeType.paragraph, true))
#
#             para.paragraph_format.bidi = true
#
#
#         # If we enable a right-to-left editing language in Microsoft Word, our fields will display numbers.
#         # Otherwise, they will display "###".
#         doc.save(aeb.artifacts_dir + "Field.bidioutline.docx")
#         #ExEnd
#
#         doc = new Document(aeb.artifacts_dir + "Field.bidioutline.docx")
#
#         foreach (Field fieldBidiOutline in doc.range.fields)
#             TestUtil.verify_field(FieldType.field_bidi_outline, " BIDIOUTLINE ", string.empty, fieldBidiOutline)
#
#
#     def test_legacy(self) :
#
#         #ExStart
#         #ExFor:FieldEmbed
#         #ExFor:FieldShape
#         #ExFor:FieldShape.text
#         #ExSummary:Shows how some older Microsoft Word fields such as SHAPE and EMBED are handled during loading.
#         # Open a document that was created in Microsoft Word 2003.
#         Document doc = new Document(aeb.my_dir + "Legacy fields.doc")
#
#         # If we open the Word document and press Alt+F9, we will see a SHAPE and an EMBED field.
#         # A SHAPE field is the anchor/canvas for an AutoShape object with the "In line with text" wrapping style enabled.
#         # An EMBED field has the same function, but for an embedded object,
#         # such as a spreadsheet from an external Excel document.
#         # However, these fields will not appear in the document's Fields collection.
#         self.assertEqual(0, doc.range.fields.count)
#
#         # These fields are supported only by old versions of Microsoft Word.
#         # The document loading process will convert these fields into Shape objects,
#         # which we can access in the document's node collection.
#         NodeCollection shapes = doc.get_child_nodes(NodeType.shape, true)
#         self.assertEqual(3, shapes.count)
#
#         # The first Shape node corresponds to the SHAPE field in the input document,
#         # which is the inline canvas for the AutoShape.
#         Shape shape = (Shape)shapes[0]
#         self.assertEqual(ShapeType.image, shape.shape_type)
#
#         # The second Shape node is the AutoShape itself.
#         shape = (Shape)shapes[1]
#         self.assertEqual(ShapeType.can, shape.shape_type)
#
#         # The third Shape is what was the EMBED field that contained the external spreadsheet.
#         shape = (Shape)shapes[2]
#         self.assertEqual(ShapeType.ole_object, shape.shape_type)
#         #ExEnd
#
#
#     def test_set_field_index_format(self) :
#
#         #ExStart
#         #ExFor:FieldOptions.field_index_format
#         #ExSummary:Shows how to formatting FieldIndex fields.
#         doc = aw.Document()
#         builder = aw.DocumentBuilder(doc)
#         builder.write("A")
#         builder.insert_break(BreakType.line_break)
#         builder.insert_field("XE \"A\"")
#         builder.write("B")
#
#         builder.insert_field(" INDEX \\e \" · \" \\h \"A\" \\c \"2\" \\z \"1033\"", null)
#
#         doc.field_options.field_index_format = FieldIndexFormat.fancy
#         doc.update_fields()
#
#         doc.save(aeb.artifacts_dir + "Field.set_field_index_format.docx")
#         #ExEnd
#
#
#     #ExStart
#     #ExFor:ComparisonEvaluationResult.#ctor(bool)
#     #ExFor:ComparisonEvaluationResult.#ctor(string)
#     #ExFor:ComparisonEvaluationResult
#     #ExFor:ComparisonExpression
#     #ExFor:ComparisonExpression.left_expression
#     #ExFor:ComparisonExpression.comparison_operator
#     #ExFor:ComparisonExpression.right_expression
#     #ExFor:FieldOptions.comparison_expression_evaluator
#     #ExSummary:Shows how to implement custom evaluation for the IF and COMPARE fields.
#     [TestCase(" IF 0 1 2 \"true argument\" \"false argument\" ", 1, null, "true argument")] #ExSkip
#     [TestCase(" IF 0 1 2 \"true argument\" \"false argument\" ", 0, null, "false argument")] #ExSkip
#     [TestCase(" IF 0 1 2 \"true argument\" \"false argument\" ", -1, "Custom Error", "Custom Error")] #ExSkip
#     [TestCase(" IF 0 1 2 \"true argument\" \"false argument\" ", -1, null, "true argument")] #ExSkip
#     [TestCase(" COMPARE 0 1 2 ", 1, null, "1")] #ExSkip
#     [TestCase(" COMPARE 0 1 2 ", 0, null, "0")] #ExSkip
#     [TestCase(" COMPARE 0 1 2 ", -1, "Custom Error", "Custom Error")] #ExSkip
#     [TestCase(" COMPARE 0 1 2 ", -1, null, "1")] #ExSkip
#     public void ConditionEvaluationExtensionPoint(string fieldCode, sbyte comparisonResult, string comparisonError,
#         string expectedResult)
#
#         const string left = "\"left expression\""
#         const string @operator = "<>"
#         const string right = "\"right expression\""
#
#         DocumentBuilder builder = new DocumentBuilder()
#
#         # Field codes that we use in this example:
#         # 1.   " IF 0 1 2 \"true argument\" \"false argument\" ".
#         # 2.   " COMPARE 0 1 2 ".
#         Field field = builder.insert_field(string.format(fieldCode, left, @operator, right), null)
#
#         # If the "comparisonResult" is undefined, we create "ComparisonEvaluationResult" with string, instead of bool.
#         ComparisonEvaluationResult result = comparisonResult != -1
#             ? new ComparisonEvaluationResult(comparisonResult == 1)
#             : comparisonError != null ? new ComparisonEvaluationResult(comparisonError) : null
#
#         ComparisonExpressionEvaluator evaluator = new ComparisonExpressionEvaluator(result)
#         builder.document.field_options.comparison_expression_evaluator = evaluator
#
#         builder.document.update_fields()
#
#         self.assertEqual(expectedResult, field.result)
#         evaluator.assert_invocations_count(1).assert_invocation_arguments(0, left, @operator, right)
#
#
#     # <summary>
#     # Comparison expressions evaluation for the FieldIf and FieldCompare.
#     # </summary>
#     private class ComparisonExpressionEvaluator : IComparisonExpressionEvaluator
#
#         public ComparisonExpressionEvaluator(ComparisonEvaluationResult result)
#
#             mResult = result
#
#
#         public ComparisonEvaluationResult Evaluate(Field field, ComparisonExpression expression)
#
#             mInvocations.add(new[]
#
#                 expression.left_expression,
#                 expression.comparison_operator,
#                 expression.right_expression
#             )
#
#             return mResult
#
#
#         public ComparisonExpressionEvaluator AssertInvocationsCount(int expected)
#
#             self.assertEqual(expected, mInvocations.count)
#             return this
#
#
#         public ComparisonExpressionEvaluator AssertInvocationArguments(
#             int invocationIndex,
#             string expectedLeftExpression,
#             string expectedComparisonOperator,
#             string expectedRightExpression)
#
#             string[] arguments = mInvocations[invocationIndex]
#
#             self.assertEqual(expectedLeftExpression, arguments[0])
#             self.assertEqual(expectedComparisonOperator, arguments[1])
#             self.assertEqual(expectedRightExpression, arguments[2])
#
#             return this
#
#
#         private readonly ComparisonEvaluationResult mResult
#         private readonly List<string[]> mInvocations = new List<string[]>()
#
#     #ExEnd
#
#     def test_comparison_expression_evaluator_nested_fields(self) :
#
#         Document document = new Document()
#
#         new FieldBuilder(FieldType.field_if)
#             .add_argument(
#                 new FieldBuilder(FieldType.field_if)
#                     .add_argument(123)
#                     .add_argument(">")
#                     .add_argument(666)
#                     .add_argument("left greater than right")
#                     .add_argument("left less than right"))
#             .add_argument("<>")
#             .add_argument(new FieldBuilder(FieldType.field_if)
#                 .add_argument("left expression")
#                 .add_argument("=")
#                 .add_argument("right expression")
#                 .add_argument("expression are equal")
#                 .add_argument("expression are not equal"))
#             .add_argument(new FieldBuilder(FieldType.field_if)
#                     .add_argument(new FieldArgumentBuilder()
#                         .add_text("#")
#                         .add_field(new FieldBuilder(FieldType.field_page)))
#                     .add_argument("=")
#                     .add_argument(new FieldArgumentBuilder()
#                         .add_text("#")
#                         .add_field(new FieldBuilder(FieldType.field_num_pages)))
#                     .add_argument("the last page")
#                     .add_argument("not the last page"))
#             .add_argument(new FieldBuilder(FieldType.field_if)
#                     .add_argument("unexpected")
#                     .add_argument("=")
#                     .add_argument("unexpected")
#                     .add_argument("unexpected")
#                     .add_argument("unexpected"))
#             .build_and_insert(document.first_section.body.first_paragraph)
#
#         ComparisonExpressionEvaluator evaluator = new ComparisonExpressionEvaluator(null)
#         document.field_options.comparison_expression_evaluator = evaluator
#
#         document.update_fields()
#
#         evaluator
#             .assert_invocations_count(4)
#             .assert_invocation_arguments(0, "123", ">", "666")
#             .assert_invocation_arguments(1, "\"left expression\"", "=", "\"right expression\"")
#             .assert_invocation_arguments(2, "left less than right", "<>", "expression are not equal")
#             .assert_invocation_arguments(3, "\"#1\"", "=", "\"#1\"")
#
#
#     def test_comparison_expression_evaluator_header_footer_fields(self) :
#
#         Document document = new Document()
#         DocumentBuilder builder = new DocumentBuilder(document)
#
#         builder.insert_break(BreakType.page_break)
#         builder.insert_break(BreakType.page_break)
#         builder.move_to_header_footer(HeaderFooterType.header_primary)
#
#         new FieldBuilder(FieldType.field_if)
#             .add_argument(new FieldBuilder(FieldType.field_page))
#             .add_argument("=")
#             .add_argument(new FieldBuilder(FieldType.field_num_pages))
#             .add_argument(new FieldArgumentBuilder()
#                 .add_field(new FieldBuilder(FieldType.field_page))
#                 .add_text(" / ")
#                 .add_field(new FieldBuilder(FieldType.field_num_pages)))
#             .add_argument(new FieldArgumentBuilder()
#                 .add_field(new FieldBuilder(FieldType.field_page))
#                 .add_text(" / ")
#                 .add_field(new FieldBuilder(FieldType.field_num_pages)))
#             .build_and_insert(builder.current_paragraph)
#
#         ComparisonExpressionEvaluator evaluator = new ComparisonExpressionEvaluator(null)
#         document.field_options.comparison_expression_evaluator = evaluator
#
#         document.update_fields()
#
#         evaluator
#             .assert_invocations_count(3)
#             .assert_invocation_arguments(0, "1", "=", "3")
#             .assert_invocation_arguments(1, "2", "=", "3")
#             .assert_invocation_arguments(2, "3", "=", "3")
#


if __name__ == '__main__':
    unittest.main()
