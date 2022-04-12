# Copyright (c) 2001-2022 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import unittest
import io
from datetime import datetime
from enum import Enum

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, IMAGE_DIR
from document_helper import DocumentHelper

class ExField(ApiExampleBase):

    def test_get_field_from_document(self):

        #ExStart
        #ExFor:FieldType
        #ExFor:FieldChar
        #ExFor:FieldChar.field_type
        #ExFor:FieldChar.is_dirty
        #ExFor:FieldChar.is_locked
        #ExFor:FieldChar.get_field
        #ExFor:Field.is_locked
        #ExSummary:Shows how to work with a FieldStart node.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        field = builder.insert_field(aw.fields.FieldType.FIELD_DATE, True).as_field_date()
        field.format.date_time_format = "dddd, MMMM dd, yyyy"
        field.update()

        field_start = field.start

        self.assertEqual(aw.fields.FieldType.FIELD_DATE, field_start.field_type)
        self.assertFalse(field_start.is_dirty)
        self.assertFalse(field_start.is_locked)

        # Retrieve the facade object which represents the field in the document.
        field = field_start.get_field().as_field_date()

        self.assertFalse(field.is_locked)
        self.assertEqual(" DATE  \\@ \"dddd, MMMM dd, yyyy\"", field.get_field_code())

        # Update the field to show the current date.
        field.update()
        #ExEnd

        doc = DocumentHelper.save_open(doc)

        self.verify_field(aw.fields.FieldType.FIELD_DATE, " DATE  \\@ \"dddd, MMMM dd, yyyy\"", datetime.now().strftime("%A, %B %d, %Y"), doc.range.fields[0])

    def test_get_field_code(self):

        #ExStart
        #ExFor:Field.get_field_code()
        #ExFor:Field.get_field_code(bool)
        #ExSummary:Shows how to get a field's field code.
        # Open a document which contains a MERGEFIELD inside an IF field.
        doc = aw.Document(MY_DIR + "Nested fields.docx")
        field_if = doc.range.fields[0].as_field_if()

        # There are two ways of getting a field's field code:
        # 1 -  Omit its inner fields:
        self.assertEqual(" IF  > 0 \" (surplus of ) \" \"\" ", field_if.get_field_code(False))

        # 2 -  Include its inner fields:
        self.assertEqual(" IF \u0013 MERGEFIELD NetIncome \u0014\u0015 > 0 \" (surplus of \u0013 MERGEFIELD  NetIncome \\f $ \u0014\u0015) \" \"\" ",
            field_if.get_field_code(True))

        # By default, the "get_field_code" method displays inner fields.
        self.assertEqual(field_if.get_field_code(), field_if.get_field_code(True))
        #ExEnd

    def test_display_result(self):

        #ExStart
        #ExFor:Field.display_result
        #ExSummary:Shows how to get the real text that a field displays in the document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("This document was written by ")
        field_author = builder.insert_field(aw.fields.FieldType.FIELD_AUTHOR, True).as_field_author()
        field_author.author_name = "John Doe"

        # We can use the "display_result" property to verify what exact text
        # a field would display in its place in the document.
        self.assertEqual("", field_author.display_result)

        # Fields do not maintain accurate result values in real-time.
        # To make sure our fields display accurate results at any given time,
        # such as right before a save operation, we need to update them manually.
        field_author.update()

        self.assertEqual("John Doe", field_author.display_result)

        doc.save(ARTIFACTS_DIR + "Field.display_result.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.display_result.docx")

        self.assertEqual("John Doe", doc.range.fields[0].display_result)

    def test_create_with_field_builder(self):

        #ExStart
        #ExFor:FieldBuilder.__init__(FieldType)
        #ExFor:FieldBuilder.build_and_insert(Inline)
        #ExSummary:Shows how to create and insert a field using a field builder.
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
        doc.save(ARTIFACTS_DIR + "Field.create_with_field_builder.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.create_with_field_builder.docx")

        self.verify_field(aw.fields.FieldType.FIELD_BARCODE, " BARCODE 90210 \\f A \\u ", "", doc.range.fields[0])

        self.assertEqual(doc.first_section.body.first_paragraph.runs[11].previous_sibling, doc.range.fields[0].end)
        self.assertEqual(f"{aw.ControlChar.FIELD_START_CHAR} BARCODE 90210 \\f A \\u {aw.ControlChar.FIELD_END_CHAR} Hello world! This text is one Run, which is an inline node.",
            doc.get_text().strip())

    def test_rev_num(self):

        #ExStart
        #ExFor:BuiltInDocumentProperties.revision_number
        #ExFor:FieldRevNum
        #ExSummary:Shows how to work with REVNUM fields.
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
        self.assertEqual("1", field.result) #ExSkip
        field.update()

        self.assertEqual("2", field.result)
        #ExEnd

        doc = DocumentHelper.save_open(doc)
        self.assertEqual(2, doc.built_in_document_properties.revision_number)

        self.verify_field(aw.fields.FieldType.FIELD_REVISION_NUM, " REVNUM ", "2", doc.range.fields[0])

    def test_insert_field_none(self):

        #ExStart
        #ExFor:FieldUnknown
        #ExSummary:Shows how to work with 'FieldNone' field in a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a field that does not denote an objective field type in its field code.
        field = builder.insert_field(" NOTAREALFIELD //a")

        # The "FieldNone" field type is reserved for fields such as these.
        self.assertEqual(aw.fields.FieldType.FIELD_NONE, field.type)

        # We can also still work with these fields and assign them as instances of the FieldUnknown class.
        field_unknown = field.as_field_unknown()
        self.assertEqual(" NOTAREALFIELD //a", field_unknown.get_field_code())
        #ExEnd

        doc = DocumentHelper.save_open(doc)

        self.verify_field(aw.fields.FieldType.FIELD_NONE, " NOTAREALFIELD //a", "Error! Bookmark not defined.", doc.range.fields[0])

    def test_insert_tc_field(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a TC field at the current document builder position.
        builder.insert_field("TC \"Entry Text\" \\f t")

    #def test_insert_tc_fields_at_text(self):

    #    doc = aw.Document()

    #    options = aw.replacing.FindReplaceOptions()
    #    options.replacing_callback = ExField.InsertTcFieldHandler("Chapter 1", "\\l 1")

    #    # Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
    #    doc.range.replace_regeg("The Beginning", "", options)

    #class InsertTcFieldHandler(aw.replacing.IReplacingCallback):

    #    def __init__(self, text: str, switches: str):
    #        """The display text and switches to use for each TC field. Display name can be an empty String or null."""

    #        # Store the text and switches to be used for the TC fields.
    #        self.field_text = text
    #        self.field_switches = switches

    #    def replacing(self, args: aw.replacing.ReplacingArgs) -> aw.replacing.ReplaceAction:

    #        builder = aw.DocumentBuilder(args.match_node.document.as_document())
    #        builder.move_to(args.match_node)

    #        # If the user-specified text is used in the field as display text, use that, otherwise
    #        # use the match String as the display text.
    #        insert_text = self.field_text if self.field_text else args.match.value

    #        # Insert the TC field before this node using the specified String
    #        # as the display text and user-defined switches.
    #        builder.insert_field(f"TC \"{insert_text}\" {self.field_switches}")

    #        return aw.replacing.ReplaceAction.SKIP

    def test_field_locale(self):

        #ExStart
        #ExFor:Field.locale_id
        #ExSummary:Shows how to insert a field and work with its locale.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a DATE field, and then print the date it will display.
        # Your thread's current culture determines the formatting of the date.
        field = builder.insert_field("DATE")
        print(f"Today's date, as displayed in the \"{CultureInfo.current_culture.english_name}\" culture: {field.result}")

        self.assertEqual(1033, field.locale_id)
        self.assertEqual(aw.fields.FieldUpdateCultureSource.CURRENT_THREAD, doc.field_options.field_update_culture_source) #ExSkip

        # Changing the culture of our thread will impact the result of the DATE field.
        # Another way to get the DATE field to display a date in a different culture is to use its LocaleId property.
        # This way allows us to avoid changing the thread's culture to get this effect.
        doc.field_options.field_update_culture_source = aw.fields.FieldUpdateCultureSource.FIELD_CODE
        de_culture = CultureInfo("de-DE")
        field.locale_id = de_culture.LCID
        field.update()

        print(f"Today's date, as displayed according to the \"{CultureInfo.get_culture_info(field.LocaleId).english_name}\" culture: {field.Result}")
        #ExEnd

        doc = DocumentHelper.save_open(doc)
        field = doc.range.fields[0]

        self.verify_field(aw.fields.FieldType.FIELD_DATE, "DATE", datetime.now.to_string(de.date_time_format.short_date_pattern), field)
        self.assertEqual(CultureInfo("de-DE").lcid, field.locale_id)

    @unittest.skip("WORDSNET-16037")
    def test_update_dirty_fields(self):

        for update_dirty_fields in (True, False):
            with self.subTest(update_dirty_fields=update_dirty_fields):
                #ExStart
                #ExFor:Field.is_dirty
                #ExFor:LoadOptions.update_dirty_fields
                #ExSummary:Shows how to use special property for updating field result.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Give the document's built-in "author" property value, and then display it with a field.
                doc.built_in_document_properties.author = "John Doe"
                field = builder.insert_field(aw.fields.FieldType.FIELD_AUTHOR, True).as_field_author()

                self.assertFalse(field.is_dirty)
                self.assertEqual("John Doe", field.result)

                # Update the property. The field still displays the old value.
                doc.built_in_document_properties.author = "John & Jane Doe"

                self.assertEqual("John Doe", field.result)

                # Since the field's value is out of date, we can mark it as "dirty".
                # This value will stay out of date until we update the field manually with the Field.update() method.
                field.is_dirty = True

                with io.BytesIO() as doc_stream:
                    # If we save without calling an update method,
                    # the field will keep displaying the out of date value in the output document.
                    doc.save(doc_stream, aw.SaveFormat.DOCX)

                    # The LoadOptions object has an option to update all fields
                    # marked as "dirty" when loading the document.
                    options = aw.loading.LoadOptions()
                    options.update_dirty_fields = update_dirty_fields
                    doc = aw.Document(doc_stream, options)

                    self.assertEqual("John & Jane Doe", doc.built_in_document_properties.author)

                    field = doc.range.fields[0].as_field_author()

                    # Updating dirty fields like this automatically set their "is_dirty" flag to False.
                    if update_dirty_fields:
                        self.assertEqual("John & Jane Doe", field.result)
                        self.assertFalse(field.is_dirty)
                    else:
                        self.assertEqual("John Doe", field.result)
                        self.assertTrue(field.is_dirty)

                #ExEnd

    def test_insert_field_with_field_builder_exception(self):

        doc = aw.Document()

        run = DocumentHelper.insert_new_run(doc, " Hello World!", 0)

        argument_builder = aw.fields.FieldArgumentBuilder()
        argument_builder.add_field(aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_MERGE_FIELD))
        argument_builder.add_node(run)
        argument_builder.add_text("Text argument builder")

        field_builder = aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_INCLUDE_TEXT)

        with self.assertRaises(Exception):
            field_builder.add_argument(argument_builder).add_argument("=").add_argument("BestField").add_argument(10).add_argument(20.0).build_and_insert(run)

    #def test_bar_code_word2_pdf(self):

    #    doc = aw.Document(MY_DIR + "Field sample - BARCODE.docx")

    #    doc.field_options.barcode_generator = CustomBarcodeGenerator()

    #    doc.save(ARTIFACTS_DIR + "Field.bar_code_word2_pdf.pdf")

    #    with ExField.bar_code_reader_pdf(ARTIFACTS_DIR + "Field.bar_code_word2_pdf.pdf") as bar_code_reader:

    #        self.assertEqual("QR", bar_code_reader.found_bar_codes[0].code_type_name)

    #def bar_code_reader_pdf(filename: str) -> aspose.barcode.barcoderecognition.BarCodeReader:

    #    # Set license for Aspose.BarCode.
    #    licence_bar_code = aspose.barcode.License()
    #    licence_bar_code.set_license(LICENSE_DIR + "Aspose.total.n_e_t.lic")

    #    pdf_extractor = aspose.pdf.facades.PdfExtractor()
    #    pdf_extractor.bind_pdf(filename)

    #    # Set page range for image extraction.
    #    pdf_extractor.start_page = 1
    #    pdf_extractor.end_page = 1

    #    pdf_extractor.extract_image()

    #    image_stream = io.BytesIO()
    #    pdf_extractor.get_next_image(image_stream)
    #    image_stream.position = 0

    #    # Recognize the barcode from the image stream above.
    #    barcode_reader = aspose.barcode.barcoderecognition.BarCodeReader(image_stream, aspose.barcode.barcoderecognition.DecodeType.QR)

    #    for result in barcode_reader.read_bar_codes():
    #        print("Codetext found: " + result.code_text + ", Symbology: " + result.code_type_name)

    #    return barcode_reader

    #@unittest.skip("WORDSNET-13854")
    #def test_field_database(self):

    #    #ExStart
    #    #ExFor:FieldDatabase
    #    #ExFor:FieldDatabase.connection
    #    #ExFor:FieldDatabase.file_name
    #    #ExFor:FieldDatabase.first_record
    #    #ExFor:FieldDatabase.format_attributes
    #    #ExFor:FieldDatabase.insert_headings
    #    #ExFor:FieldDatabase.insert_once_on_mail_merge
    #    #ExFor:FieldDatabase.last_record
    #    #ExFor:FieldDatabase.query
    #    #ExFor:FieldDatabase.table_format
    #    #ExSummary:Shows how to extract data from a database and insert it as a field into a document.
    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    # This DATABASE field will run a query on a database, and display the result in a table.
    #    field = builder.insert_field(aw.fields.FieldType.FIELD_DATABASE, True).as_field_database()
    #    field.file_name = MY_DIR + r"Database\Northwind.mdb"
    #    field.connection = "DSN=MS Access Databases"
    #    field.query = "SELECT * FROM [Products]"

    #    database_path = DATABASE_DIR.replace("\\", "\\\\") + "Northwind.mdb"
    #    self.assertEqual(' DATABASE  \\d "' + database_path + '" \\c "DSN=MS Access Databases" \\s "SELECT * FROM [Products]"',
    #        field.get_field_code())

    #    # Insert another DATABASE field with a more complex query that sorts all products in descending order by gross sales.
    #    field = builder.insert_field(aw.fields.FieldType.FIELD_DATABASE, True).as_field_database()
    #    field.file_name = MY_DIR + r"Database\Northwind.mdb"
    #    field.connection = "DSN=MS Access Databases"
    #    field.query =(
    #        "SELECT [Products].ProductName, FORMAT(SUM([Order Details].UnitPrice * (1 - [Order Details].Discount) * [Order Details].Quantity), 'Currency') AS GrossSales " +
    #        "FROM([Products] " +
    #        "LEFT JOIN[Order Details] ON[Products].[ProductID] = [Order Details].[ProductID]) " +
    #        "GROUP BY[Products].ProductName " +
    #        "ORDER BY SUM([Order Details].UnitPrice* (1 - [Order Details].Discount) * [Order Details].Quantity) DESC")

    #    # These properties have the same function as LIMIT and TOP clauses.
    #    # Configure them to display only rows 1 to 10 of the query result in the field's table.
    #    field.first_record = "1"
    #    field.last_record = "10"

    #    # This property is the index of the format we want to use for our table. The list of table formats is in the "Table AutoFormat..." menu
    #    # that shows up when we create a DATABASE field in Microsoft Word. Index #10 corresponds to the "Colorful 3" format.
    #    field.table_format = "10"

    #    # The format_attributes property is a string representation of an integer which stores multiple flags.
    #    # We can patrially apply the format which the TableFormat property points to by setting different flags in this property.
    #    # The number we use is the sum of a combination of values corresponding to different aspects of the table style.
    #    # 63 represents 1 (borders) + 2 (shading) + 4 (font) + 8 (color) + 16 (autofit) + 32 (heading rows).
    #    field.format_attributes = "63"
    #    field.insert_headings = True
    #    field.insert_once_on_mail_merge = True

    #    doc.update_fields()
    #    doc.save(ARTIFACTS_DIR + "Field.database.docx")
    #    #ExEnd

    #    doc = aw.Document(ARTIFACTS_DIR + "Field.database.docx")

    #    self.assertEqual(2, doc.range.fields.count)

    #    table = doc.first_section.body.tables[0]

    #    self.assertEqual(77, table.rows.count)
    #    self.assertEqual(10, table.rows[0].cells.count)

    #    field = doc.range.fields[0].as_field_database()

    #    self.assertEqual(' DATABASE  \\d "' + database_path + '" \\c "DSN=MS Access Databases" \\s "SELECT * FROM [Products]"',
    #        field.get_field_code())

    #    self.table_matches_query_result(table, DATABASE_DIR + "Northwind.mdb", field.query)

    #    table = doc.get_child(aw.NodeType.TABLE, 1, True).as_table()
    #    field = doc.range.fields[1].as_field_database()

    #    self.assertEqual(11, table.rows.count)
    #    self.assertEqual(2, table.rows[0].cells.count)
    #    self.assertEqual("ProductName\a", table.rows[0].cells[0].get_text())
    #    self.assertEqual("GrossSales\a", table.rows[0].cells[1].get_text())

    #    self.assertEqual((
    #        ' DATABASE  \\d "' + database_path + '" \\c "DSN=MS Access Databases" ' +
    #        '\\s "SELECT [Products].ProductName, FORMAT(SUM([Order Details].UnitPrice * (1 - [Order Details].Discount) * [Order Details].Quantity), "Currency") AS GrossSales ' +
    #        'FROM([Products] ' +
    #        'LEFT JOIN[Order Details] ON[Products].[ProductID] = [Order Details].[ProductID]) ' +
    #        'GROUP BY[Products].ProductName ' +
    #        'ORDER BY SUM([Order Details].UnitPrice* (1 - [Order Details].Discount) * [Order Details].Quantity) DESC\" \\f 1 \\t 10 \\l 10 \\b 63 \\h \\o'),
    #        field.get_field_code())

    #    table.rows[0].remove()

    #    self.table_matches_query_result(table, DATABASE_DIR + "Northwind.mdb", field.query.insert(7, " TOP 10 "))

    def test_preserve_include_picture(self):

        for preserve_include_picture_field in (False, True):
            with self.subTest(preserve_include_picture_field=preserve_include_picture_field):
                #ExStart
                #ExFor:Field.update(bool)
                #ExFor:LoadOptions.preserve_include_picture_field
                #ExSummary:Shows how to preserve or discard INCLUDEPICTURE fields when loading a document.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                include_picture = builder.insert_field(aw.fields.FieldType.FIELD_INCLUDE_PICTURE, True).as_field_include_picture()
                include_picture.source_full_name = IMAGE_DIR + "Transparent background logo.png"
                include_picture.update(True)

                with io.BytesIO() as doc_stream:

                    doc.save(doc_stream, aw.saving.OoxmlSaveOptions(aw.SaveFormat.DOCX))

                    # We can set a flag in a LoadOptions object to decide whether to convert all INCLUDEPICTURE fields
                    # into image shapes when loading a document that contains them.
                    load_options = aw.loading.LoadOptions()
                    load_options.preserve_include_picture_field = preserve_include_picture_field

                    doc = aw.Document(doc_stream, load_options)

                    if preserve_include_picture_field:
                        self.assertTrue(any(f for f in doc.range.fields if f.type == aw.fields.FieldType.FIELD_INCLUDE_PICTURE))

                        doc.update_fields()
                        doc.save(ARTIFACTS_DIR + "Field.preserve_include_picture.docx")
                    else:
                        self.assertFalse(any(f for f in doc.range.fields if f.type == aw.fields.FieldType.FIELD_INCLUDE_PICTURE))

                #ExEnd

    def test_field_format(self):

        #ExStart
        #ExFor:Field.format
        #ExFor:Field.update()
        #ExFor:FieldFormat
        #ExFor:FieldFormat.date_time_format
        #ExFor:FieldFormat.numeric_format
        #ExFor:FieldFormat.general_formats
        #ExFor:GeneralFormat
        #ExFor:GeneralFormatCollection
        #ExFor:GeneralFormatCollection.add(GeneralFormat)
        #ExFor:GeneralFormatCollection.count
        #ExFor:GeneralFormatCollection.__getitem__(int)
        #ExFor:GeneralFormatCollection.remove(GeneralFormat)
        #ExFor:GeneralFormatCollection.remove_at(int)
        #ExFor:GeneralFormatCollection.__iter__
        #ExSummary:Shows how to format field results.
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
        print(f"Today's date, in {format.date_time_format} format:\n\t{field.result}")

        # 3 -  General format:
        field = builder.insert_field("= 25 + 33")
        format = field.format
        format.general_formats.add(aw.fields.GeneralFormat.LOWERCASE_ROMAN)
        format.general_formats.add(aw.fields.GeneralFormat.UPPER)
        field.update()

        for index, general_format in enumerate(format.general_formats):
            print(f"General format index {index}: {general_format}")

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
        #ExEnd

    def test_unlink(self):

        #ExStart
        #ExFor:Document.unlink_fields
        #ExSummary:Shows how to unlink all fields in the document.
        doc = aw.Document(MY_DIR + "Linked fields.docx")

        doc.unlink_fields()
        #ExEnd

        doc = DocumentHelper.save_open(doc)
        para_with_fields = DocumentHelper.get_paragraph_text(doc, 0)

        self.assertEqual("Fields.Docx   Элементы указателя не найдены.     1.\r", para_with_fields)

    def test_unlink_all_fields_in_range(self):

        #ExStart
        #ExFor:Range.unlink_fields
        #ExSummary:Shows how to unlink all fields in a range.
        doc = aw.Document(MY_DIR + "Linked fields.docx")

        new_section = doc.sections[0].clone(True).as_section()
        doc.sections.add(new_section)

        doc.sections[1].range.unlink_fields()
        #ExEnd

        doc = DocumentHelper.save_open(doc)
        sec_with_fields = DocumentHelper.get_section_text(doc, 1)

        self.assertTrue(sec_with_fields.strip().endswith(
            "Fields.Docx   Элементы указателя не найдены.     3.\rОшибка! Не указана последовательность.    Fields.Docx   Элементы указателя не найдены.     4."))

    def test_unlink_single_field(self):

        #ExStart
        #ExFor:Field.unlink
        #ExSummary:Shows how to unlink a field.
        doc = aw.Document(MY_DIR + "Linked fields.docx")
        doc.range.fields[1].unlink()
        #ExEnd

        doc = DocumentHelper.save_open(doc)
        para_with_fields = DocumentHelper.get_paragraph_text(doc, 0)

        self.assertTrue(para_with_fields.strip().endswith(
            "FILENAME  \\* Caps  \\* MERGEFORMAT \u0014Fields.Docx\u0015   Элементы указателя не найдены.     \u0013 LISTNUM  LegalDefault \u0015"))

    def test_update_toc_page_numbers(self):

        doc = aw.Document(MY_DIR + "Field sample - TOC.docx")

        start_node = DocumentHelper.get_paragraph(doc, 2)
        end_node = None

        paragraph_collection = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)

        for para in paragraph_collection:
            for run in para.as_paragraph().runs:
                run = run.as_run()
                if aw.ControlChar.PAGE_BREAK in run.text:
                    end_node = run
                    break

        if start_node is not None and end_node is not None:
            ExField.remove_sequence(start_node, end_node)

            start_node.remove()
            end_node.remove()

        f_start = doc.get_child_nodes(aw.NodeType.FIELD_START, True)

        for field in f_start:
            field = field.as_field_start()
            f_type = field.field_type
            if f_type == aw.fields.FieldType.FIELD_TOC:
                para = field.get_ancestor(aw.NodeType.PARAGRAPH).as_paragraph()
                para.range.update_fields()
                break

        doc.save(ARTIFACTS_DIR + "Field.update_toc_page_numbers.docx")

    @staticmethod
    def remove_sequence(start: aw.Node, end: aw.Node):

        cur_node = start.next_pre_order(start.document)
        while cur_node is not None and cur_node != end:
            next_node = cur_node.next_pre_order(start.document)
            if cur_node.is_composite:
                cur_composite = cur_node.as_composite_node()
                if (not cur_composite.get_child_nodes(aw.NodeType.ANY, True).contains(end) and
                    not cur_composite.get_child_nodes(aw.NodeType.ANY, True).contains(start)):

                    next_node = cur_node.next_sibling
                    cur_node.remove()
            else:
                cur_node.remove()

            cur_node = next_node

    ##ExStart
    ##ExFor:Fields.field_ask
    ##ExFor:Fields.FieldAsk.bookmark_name
    ##ExFor:Fields.FieldAsk.default_response
    ##ExFor:Fields.FieldAsk.prompt_once_on_mail_merge
    ##ExFor:Fields.FieldAsk.prompt_text
    ##ExFor:FieldOptions.user_prompt_respondent
    ##ExFor:IFieldUserPromptRespondent
    ##ExFor:IFieldUserPromptRespondent.respond(str,str)
    ##ExSummary:Shows how to create an ASK field, and set its properties.
    #def test_field_ask(self):

    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    # Place a field where the response to our ASK field will be placed.
    #    field_ref = builder.insert_field(aw.fields.FieldType.FIELD_REF, True).as_field_ref()
    #    field_ref.bookmark_name = "MyAskField"
    #    builder.writeln()

    #    self.assertEqual(" REF  MyAskField", field_ref.get_field_code())

    #    # Insert the ASK field and edit its properties to reference our REF field by bookmark name.
    #    field_ask = builder.insert_field(aw.fields.FieldType.FIELD_ASK, True).as_field_ask()
    #    field_ask.bookmark_name = "MyAskField"
    #    field_ask.prompt_text = "Please provide a response for this ASK field"
    #    field_ask.default_response = "Response from within the field."
    #    field_ask.prompt_once_on_mail_merge = True
    #    builder.writeln()

    #    self.assertEqual(
    #        ' ASK  MyAskField "Please provide a response for this ASK field" \\d "Response from within the field." \\o',
    #        field_ask.get_field_code())

    #    # ASK fields apply the default response to their respective REF fields during a mail merge.
    #    table = DataTable("My Table")
    #    table.columns.add("Column 1")
    #    table.rows.add("Row 1")
    #    table.rows.add("Row 2")

    #    field_merge_field = builder.insert_field(aw.fields.FieldType.FIELD_MERGE_FIELD, True).as_field_merge_field()
    #    field_merge_field.field_name = "Column 1"

    #    # We can modify or override the default response in our ASK fields with a custom prompt responder,
    #    # which will occur during a mail merge.
    #    doc.field_options.user_prompt_respondent = ExField.MyPromptRespondent()
    #    doc.mail_merge.execute(table)

    #    doc.update_fields()
    #    doc.save(ARTIFACTS_DIR + "Field.a_s_k.docx")
    #    self._test_field_ask(table, doc) #ExSkip

    #class MyPromptRespondent(aw.fields.IFieldUserPromptRespondent):
    #    """Prepends text to the default response of an ASK field during a mail merge."""

    #    def respond(self, prompt_text: str, default_response: str) -> str:

    #        return "Response from MyPromptRespondent. " + default_response

    ##ExEnd

    #def _test_field_ask(self, data_table: DataTable, doc: aw.Document):

    #    doc = DocumentHelper.save_open(doc)

    #    for field in doc.range.fields:
    #        if field.type == aw.fields.FieldType.FIELD_REF:
    #            field_ref = field.as_field_ref()
    #            break

    #    self.verify_field(aw.fields.FieldType.FIELD_REF,
    #        " REF  MyAskField", "Response from MyPromptRespondent. Response from within the field.", field_ref)

    #    for field in doc.range.fields:
    #        if field.type == aw.fields.FieldType.FIELD_ASK:
    #            field_ask = field.as_field_ask()
    #            break

    #    self.verify_field(aw.fields.FieldType.FIELD_ASK,
    #        " ASK  MyAskField \"Please provide a response for this ASK field\" \\d \"Response from within the field.\" \\o",
    #        "Response from MyPromptRespondent. Response from within the field.", field_ask)

    #    self.assertEqual("MyAskField", field_ask.bookmark_name)
    #    self.assertEqual("Please provide a response for this ASK field", field_ask.prompt_text)
    #    self.assertEqual("Response from within the field.", field_ask.default_response)
    #    self.assertEqual(True, field_ask.prompt_once_on_mail_merge)

    #    self.mail_merge_matches_data_table(data_table, doc, True)

    def test_field_advance(self):

        #ExStart
        #ExFor:FieldAdvance
        #ExFor:FieldAdvance.down_offset
        #ExFor:FieldAdvance.horizontal_position
        #ExFor:FieldAdvance.left_offset
        #ExFor:FieldAdvance.right_offset
        #ExFor:FieldAdvance.up_offset
        #ExFor:FieldAdvance.vertical_position
        #ExSummary:Shows how to insert an ADVANCE field, and edit its properties.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("This text is in its normal place.")

        # Below are two ways of using the ADVANCE field to adjust the position of text that follows it.
        # The effects of an ADVANCE field continue to be applied until the paragraph ends,
        # or another ADVANCE field updates the offset/coordinate values.
        # 1 -  Specify a directional offset:
        field = builder.insert_field(aw.fields.FieldType.FIELD_ADVANCE, True).as_field_advance()
        self.assertEqual(aw.fields.FieldType.FIELD_ADVANCE, field.type) #ExSkip
        self.assertEqual(" ADVANCE ", field.get_field_code()) #ExSkip
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

        doc.save(ARTIFACTS_DIR + "Field.field_advance.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_advance.docx")

        field = doc.range.fields[0].as_field_advance()

        self.verify_field(aw.fields.FieldType.FIELD_ADVANCE, " ADVANCE  \\r 5 \\u 5", "", field)
        self.assertEqual("5", field.right_offset)
        self.assertEqual("5", field.up_offset)

        field = doc.range.fields[1].as_field_advance()

        self.verify_field(aw.fields.FieldType.FIELD_ADVANCE, " ADVANCE  \\d 5 \\l 100", "", field)
        self.assertEqual("5", field.down_offset)
        self.assertEqual("100", field.left_offset)

        field = doc.range.fields[2].as_field_advance()

        self.verify_field(aw.fields.FieldType.FIELD_ADVANCE, " ADVANCE  \\x -100 \\y 200", "", field)
        self.assertEqual("-100", field.horizontal_position)
        self.assertEqual("200", field.vertical_position)

    def test_field_address_block(self):

        #ExStart
        #ExFor:FieldAddressBlock.excluded_country_or_region_name
        #ExFor:FieldAddressBlock.format_address_on_country_or_region
        #ExFor:FieldAddressBlock.include_country_or_region_name
        #ExFor:FieldAddressBlock.language_id
        #ExFor:FieldAddressBlock.name_and_address_format
        #ExSummary:Shows how to insert an ADDRESSBLOCK field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        field = builder.insert_field(aw.fields.FieldType.FIELD_ADDRESS_BLOCK, True).as_field_address_block()

        self.assertEqual(" ADDRESSBLOCK ", field.get_field_code())

        # Setting this to "2" will include all countries and regions,
        # unless it is the one specified in the "excluded_country_or_region_name" property.
        field.include_country_or_region_name = "2"
        field.format_address_on_country_or_region = True
        field.excluded_country_or_region_name = "United States"
        field.name_and_address_format = "<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>"

        # By default, this property will contain the language ID of the first character of the document.
        # We can set a different culture for the field to format the result with like this.
        field.language_id = "1033" # en-US

        self.assertEqual(
            " ADDRESSBLOCK  \\c 2 \\d \\e \"United States\" \\f \"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>\" \\l 1033",
            field.get_field_code())
        #ExEnd

        doc = DocumentHelper.save_open(doc)
        field = doc.range.fields[0].as_field_address_block()

        self.verify_field(aw.fields.FieldType.FIELD_ADDRESS_BLOCK,
            ' ADDRESSBLOCK  \\c 2 \\d \\e "United States" \\f "<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>" \\l 1033',
            '«AddressBlock»', field)
        self.assertEqual("2", field.include_country_or_region_name)
        self.assertEqual(True, field.format_address_on_country_or_region)
        self.assertEqual("United States", field.excluded_country_or_region_name)
        self.assertEqual("<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>",
            field.name_and_address_format)
        self.assertEqual("1033", field.language_id)

    ##ExStart
    ##ExFor:FieldCollection
    ##ExFor:FieldCollection.count
    ##ExFor:FieldCollection.__iter__
    ##ExFor:FieldStart
    ##ExFor:FieldStart.accept(DocumentVisitor)
    ##ExFor:FieldSeparator
    ##ExFor:FieldSeparator.accept(DocumentVisitor)
    ##ExFor:FieldEnd
    ##ExFor:FieldEnd.accept(DocumentVisitor)
    ##ExFor:FieldEnd.has_separator
    ##ExFor:Field.end
    ##ExFor:Field.separator
    ##ExFor:Field.start
    ##ExSummary:Shows how to work with a collection of fields.
    #def test_field_collection(self):

    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    builder.insert_field(" DATE \\@ \"dddd, d MMMM yyyy\" ")
    #    builder.insert_field(" TIME ")
    #    builder.insert_field(" REVNUM ")
    #    builder.insert_field(" AUTHOR  \"John Doe\" ")
    #    builder.insert_field(" SUBJECT \"My Subject\" ")
    #    builder.insert_field(" QUOTE \"Hello world!\" ")
    #    doc.update_fields()

    #    fields = doc.range.fields

    #    self.assertEqual(6, fields.count)

    #    # Iterate over the field collection, and print contents and type
    #    # of every field using a custom visitor implementation.
    #    field_visitor = ExField.FieldVisitor()

    #    for field in fields:
    #        field.start.accept(field_visitor)
    #        if field.separator is not None:
    #            field.separator.accept(field_visitor)
    #        field.end.accept(field_visitor)

    #    if not fields:
    #        print("There are no fields in the document.")

    #    print(field_visitor.get_text())
    #    self._test_field_collection(field_visitor.get_text()) #ExSkip

    #class FieldVisitor(aw.DocumentVisitor):
    #    """Document visitor implementation that prints field info."""

    #    def __init__(self):

    #        self.builder = io.StringIO()

    #    def get_text(self) -> str:
    #        """Gets the plain text of the document that was accumulated by the visitor."""

    #        return self.builder.getvalue()

    #    def visit_field_start(self, field_start: aw.fields.FieldStart) -> aw.VisitorAction:
    #        """Called when a FieldStart node is encountered in the document."""

    #        self.builder.append_line("Found field: " + field_start.field_type)
    #        self.builder.append_line("\tField code: " + field_start.get_field().get_field_code())
    #        self.builder.append_line("\tDisplayed as: " + field_start.get_field().result)

    #        return aw.VisitorAction.CONTINUE

    #    def visit_field_separator(self, field_separator: aw.fields.FieldSeparator) -> aw.VisitorAction:
    #        """Called when a FieldSeparator node is encountered in the document."""

    #        self.builder.append_line("\tFound separator: " + field_separator.get_text())

    #        return aw.VisitorAction.CONTINUE

    #    def VisitFieldEnd(self, field_end: aw.fields.FieldEnd) -> aw.VisitorAction:
    #        """Called when a FieldEnd node is encountered in the document."""

    #        self.builder.append_line("End of field: " + field_end.field_type)

    #        return aw.VisitorAction.CONTINUE

    ##ExEnd

    def _test_field_collection(self, field_visitor_text: str):

        self.assertIn("Found field: FieldDate", field_visitor_text)
        self.assertIn("Found field: FieldTime", field_visitor_text)
        self.assertIn("Found field: FieldRevisionNum", field_visitor_text)
        self.assertIn("Found field: FieldAuthor", field_visitor_text)
        self.assertIn("Found field: FieldSubject", field_visitor_text)
        self.assertIn("Found field: FieldQuote", field_visitor_text)

    def test_remove_fields(self):

        #ExStart
        #ExFor:FieldCollection
        #ExFor:FieldCollection.count
        #ExFor:FieldCollection.clear
        #ExFor:FieldCollection.__getitem__(int)
        #ExFor:FieldCollection.remove(Field)
        #ExFor:FieldCollection.remove_at(int)
        #ExFor:Field.remove
        #ExSummary:Shows how to remove fields from a field collection.
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
        #ExEnd

    def test_field_compare(self):

        #ExStart
        #ExFor:FieldCompare
        #ExFor:FieldCompare.comparison_operator
        #ExFor:FieldCompare.left_expression
        #ExFor:FieldCompare.right_expression
        #ExSummary:Shows how to compare expressions using a COMPARE field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        field = builder.insert_field(aw.fields.FieldType.FIELD_COMPARE, True).as_field_compare()
        field.left_expression = "3"
        field.comparison_operator = "<"
        field.right_expression = "2"
        field.update()

        # The COMPARE field displays a "0" or a "1", depending on its statement's truth.
        # The result of this statement is False so that this field will display a "0".
        self.assertEqual(" COMPARE  3 < 2", field.get_field_code())
        self.assertEqual("0", field.result)

        builder.writeln()

        field = builder.insert_field(aw.fields.FieldType.FIELD_COMPARE, True).as_field_compare()
        field.left_expression = "5"
        field.comparison_operator = "="
        field.right_expression = "2 + 3"
        field.update()

        # This field displays a "1" since the statement is True.
        self.assertEqual(" COMPARE  5 = \"2 + 3\"", field.get_field_code())
        self.assertEqual("1", field.result)

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_compare.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_compare.docx")

        field = doc.range.fields[0].as_field_compare()

        self.verify_field(aw.fields.FieldType.FIELD_COMPARE, " COMPARE  3 < 2", "0", field)
        self.assertEqual("3", field.left_expression)
        self.assertEqual("<", field.comparison_operator)
        self.assertEqual("2", field.right_expression)

        field = doc.range.fields[1].as_field_compare()

        self.verify_field(aw.fields.FieldType.FIELD_COMPARE, " COMPARE  5 = \"2 + 3\"", "1", field)
        self.assertEqual("5", field.left_expression)
        self.assertEqual("=", field.comparison_operator)
        self.assertEqual("\"2 + 3\"", field.right_expression)

    def test_field_if(self):

        #ExStart
        #ExFor:FieldIf
        #ExFor:FieldIf.comparison_operator
        #ExFor:FieldIf.evaluate_condition
        #ExFor:FieldIf.false_text
        #ExFor:FieldIf.left_expression
        #ExFor:FieldIf.right_expression
        #ExFor:FieldIf.true_text
        #ExFor:FieldIfComparisonResult
        #ExSummary:Shows how to insert an IF field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Statement 1: ")
        field = builder.insert_field(aw.fields.FieldType.FIELD_IF, True).as_field_if()
        field.left_expression = "0"
        field.comparison_operator = "="
        field.right_expression = "1"

        # The IF field will display a string from either its "true_text" property,
        # or its "false_text" property, depending on the truth of the statement that we have constructed.
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
        doc.save(ARTIFACTS_DIR + "Field.field_if.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_if.docx")
        field = doc.range.fields[0].as_field_if()

        self.verify_field(aw.fields.FieldType.FIELD_IF, " IF  0 = 1 True False", "False", field)
        self.assertEqual("0", field.left_expression)
        self.assertEqual("=", field.comparison_operator)
        self.assertEqual("1", field.right_expression)
        self.assertEqual("True", field.true_text)
        self.assertEqual("False", field.false_text)

        field = doc.range.fields[1].as_field_if()

        self.verify_field(aw.fields.FieldType.FIELD_IF, " IF  5 = \"2 + 3\" True False", "True", field)
        self.assertEqual("5", field.left_expression)
        self.assertEqual("=", field.comparison_operator)
        self.assertEqual("\"2 + 3\"", field.right_expression)
        self.assertEqual("True", field.true_text)
        self.assertEqual("False", field.false_text)

    def test_field_auto_num(self):

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

        # The separator character, which appears in the field result immediately after the number, is a full stop by default.
        # If we leave this property null, our second AUTONUM field will display "2." in the document.
        self.assertIsNone(field.separator_character)

        # We can set this property to apply the first character of its string as the new separator character.
        # In this case, our AUTONUM field will now display "2:".
        field.separator_character = ":"

        self.assertEqual(" AUTONUM  \\s :", field.get_field_code())

        doc.save(ARTIFACTS_DIR + "Field.field_auto_num.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_auto_num.docx")

        self.verify_field(aw.fields.FieldType.FIELD_AUTO_NUM, " AUTONUM ", "", doc.range.fields[0])
        self.verify_field(aw.fields.FieldType.FIELD_AUTO_NUM, " AUTONUM  \\s :", "", doc.range.fields[1])

    #ExStart
    #ExFor:FieldAutoNumLgl
    #ExFor:FieldAutoNumLgl.remove_trailing_period
    #ExFor:FieldAutoNumLgl.separator_character
    #ExSummary:Shows how to organize a document using AUTONUMLGL fields.
    def test_field_auto_num_lgl(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        filler_text = (
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
            "\nUt enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ")

        # AUTONUMLGL fields display a number that increments at each AUTONUMLGL field within its current heading level.
        # These fields maintain a separate count for each heading level,
        # and each field also displays the AUTONUMLGL field counts for all heading levels below its own.
        # Changing the count for any heading level resets the counts for all levels above that level to 1.
        # This allows us to organize our document in the form of an outline list.
        # This is the first AUTONUMLGL field at a heading level of 1, displaying "1." in the document.
        ExField._insert_numbered_clause(builder, "\tHeading 1", filler_text, aw.StyleIdentifier.HEADING1)

        # This is the second AUTONUMLGL field at a heading level of 1, so it will display "2.".
        ExField._insert_numbered_clause(builder, "\tHeading 2", filler_text, aw.StyleIdentifier.HEADING1)

        # This is the first AUTONUMLGL field at a heading level of 2,
        # and the AUTONUMLGL count for the heading level below it is "2", so it will display "2.1.".
        ExField._insert_numbered_clause(builder, "\tHeading 3", filler_text, aw.StyleIdentifier.HEADING2)

        # This is the first AUTONUMLGL field at a heading level of 3.
        # Working in the same way as the field above, it will display "2.1.1.".
        ExField._insert_numbered_clause(builder, "\tHeading 4", filler_text, aw.StyleIdentifier.HEADING3)

        # This field is at a heading level of 2, and its respective AUTONUMLGL count is at 2, so the field will display "2.2.".
        ExField._insert_numbered_clause(builder, "\tHeading 5", filler_text, aw.StyleIdentifier.HEADING2)

        # Incrementing the AUTONUMLGL count for a heading level below this one
        # has reset the count for this level so that this field will display "2.2.1.".
        ExField._insert_numbered_clause(builder, "\tHeading 6", filler_text, aw.StyleIdentifier.HEADING3)

        for field in doc.range.fields:
            if field.type == aw.fields.FieldType.FIELD_AUTO_NUM_LEGAL:
                field = field.as_field_auto_num_lgl()
                # The separator character, which appears in the field result immediately after the number,
                # is a full stop by default. If we leave this property null,
                # our last AUTONUMLGL field will display "2.2.1." in the document.
                self.assertIsNone(field.separator_character)

                # Setting a custom separator character and removing the trailing period
                # will change that field's appearance from "2.2.1." to "2:2:1".
                # We will apply this to all the fields that we have created.
                field.separator_character = ":"
                field.remove_trailing_period = True
                self.assertEqual(" AUTONUMLGL  \\s : \\e", field.get_field_code())

        doc.save(ARTIFACTS_DIR + "Field.field_auto_num_lgl.docx")
        self._test_field_auto_num_lgl(doc) #ExSkip

    @staticmethod
    def _insert_numbered_clause(builder: aw.DocumentBuilder, heading: str, contents: str, heading_style: aw.StyleIdentifier):
        """Uses a document builder to insert a clause numbered by an AUTONUMLGL field."""

        builder.insert_field(aw.fields.FieldType.FIELD_AUTO_NUM_LEGAL, True)
        builder.current_paragraph.paragraph_format.style_identifier = heading_style
        builder.writeln(heading)

        # This text will belong to the auto num legal field above it.
        # It will collapse when we click the arrow next to the corresponding AUTONUMLGL field in Microsoft Word.
        builder.current_paragraph.paragraph_format.style_identifier = aw.StyleIdentifier.BODY_TEXT
        builder.writeln(contents)

    #ExEnd

    def _test_field_auto_num_lgl(self, doc: aw.Document):

        doc = DocumentHelper.save_open(doc)

        for field in doc.range.fields:
            if field.type == aw.fields.FieldType.FIELD_AUTO_NUM_LEGAL:
                field = field.as_field_auto_num_lgl()
                self.verify_field(aw.fields.FieldType.FIELD_AUTO_NUM_LEGAL, " AUTONUMLGL  \\s : \\e", "", field)

                self.assertEqual(":", field.separator_character)
                self.assertTrue(field.remove_trailing_period)

    def test_field_auto_num_out(self):

        #ExStart
        #ExFor:FieldAutoNumOut
        #ExSummary:Shows how to number paragraphs using AUTONUMOUT fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # AUTONUMOUT fields display a number that increments at each AUTONUMOUT field.
        # Unlike AUTONUM fields, AUTONUMOUT fields use the outline numbering scheme,
        # which we can define in Microsoft Word via Format -> Bullets & Numbering -> "Outline Numbered".
        # This allows us to automatically number items like a numbered list.
        # LISTNUM fields are a newer alternative to AUTONUMOUT fields.
        # This field will display "1.".
        builder.insert_field(aw.fields.FieldType.FIELD_AUTO_NUM_OUTLINE, True)
        builder.writeln("\tParagraph 1.")

        # This field will display "2.".
        builder.insert_field(aw.fields.FieldType.FIELD_AUTO_NUM_OUTLINE, True)
        builder.writeln("\tParagraph 2.")

        for field in doc.range.fields:
            if field.type == aw.fields.FieldType.FIELD_AUTO_NUM_OUTLINE:
                field = field.as_field_auto_num_out()
                self.assertEqual(" AUTONUMOUT ", field.get_field_code())

        doc.save(ARTIFACTS_DIR + "Field.field_auto_num_out.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_auto_num_out.docx")

        for field in doc.range.fields:
            self.verify_field(aw.fields.FieldType.FIELD_AUTO_NUM_OUTLINE, " AUTONUMOUT ", "", field)

    def test_field_auto_text(self):

        #ExStart
        #ExFor:FieldAutoText
        #ExFor:FieldAutoText.entry_name
        #ExFor:FieldOptions.built_in_templates_paths
        #ExFor:FieldGlossary
        #ExFor:FieldGlossary.entry_name
        #ExSummary:Shows how to display a building block with AUTOTEXT and GLOSSARY fields.
        doc = aw.Document()

        # Create a glossary document and add an AutoText building block to it.
        doc.glossary_document = aw.buildingblocks.GlossaryDocument()
        building_block = aw.buildingblocks.BuildingBlock(doc.glossary_document)
        building_block.name = "MyBlock"
        building_block.gallery = aw.buildingblocks.BuildingBlockGallery.AUTO_TEXT
        building_block.category = "General"
        building_block.description = "MyBlock description"
        building_block.behavior = aw.buildingblocks.BuildingBlockBehavior.PARAGRAPH
        doc.glossary_document.append_child(building_block)

        # Create a source and add it as text to our building block.
        building_block_source = aw.Document()
        building_block_source_builder = aw.DocumentBuilder(building_block_source)
        building_block_source_builder.writeln("Hello World!")

        building_block_content = doc.glossary_document.import_node(building_block_source.first_section, True)
        building_block.append_child(building_block_content)

        # Set a file which contains parts that our document, or its attached template may not contain.
        doc.field_options.built_in_templates_paths = [MY_DIR + "Busniess brochure.dotx"]

        builder = aw.DocumentBuilder(doc)

        # Below are two ways to use fields to display the contents of our building block.
        # 1 -  Using an AUTOTEXT field:
        field_auto_text = builder.insert_field(aw.fields.FieldType.FIELD_AUTO_TEXT, True).as_field_auto_text()
        field_auto_text.entry_name = "MyBlock"

        self.assertEqual(" AUTOTEXT  MyBlock", field_auto_text.get_field_code())

        # 2 -  Using a GLOSSARY field:
        field_glossary = builder.insert_field(aw.fields.FieldType.FIELD_GLOSSARY, True).as_field_glossary()
        field_glossary.entry_name = "MyBlock"

        self.assertEqual(" GLOSSARY  MyBlock", field_glossary.get_field_code())

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_auto_text.glossary.dotx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_auto_text.glossary.dotx")

        self.assertEqual(0, len(doc.field_options.built_in_templates_paths))

        field_auto_text = doc.range.fields[0].as_field_auto_text()

        self.verify_field(aw.fields.FieldType.FIELD_AUTO_TEXT, " AUTOTEXT  MyBlock", "Hello World!\r", field_auto_text)
        self.assertEqual("MyBlock", field_auto_text.entry_name)

        field_glossary = doc.range.fields[1].as_field_glossary()

        self.verify_field(aw.fields.FieldType.FIELD_GLOSSARY, " GLOSSARY  MyBlock", "Hello World!\r", field_glossary)
        self.assertEqual("MyBlock", field_glossary.entry_name)

    #ExStart
    #ExFor:FieldAutoTextList
    #ExFor:FieldAutoTextList.entry_name
    #ExFor:FieldAutoTextList.list_style
    #ExFor:FieldAutoTextList.screen_tip
    #ExSummary:Shows how to use an AUTOTEXTLIST field to select from a list of AutoText entries.
    def test_field_auto_text_list(self):

        doc = aw.Document()

        # Create a glossary document and populate it with auto text entries.
        doc.glossary_document = aw.buildingblocks.GlossaryDocument()
        ExField.append_auto_text_entry(doc.glossary_document, "AutoText 1", "Contents of AutoText 1")
        ExField.append_auto_text_entry(doc.glossary_document, "AutoText 2", "Contents of AutoText 2")
        ExField.append_auto_text_entry(doc.glossary_document, "AutoText 3", "Contents of AutoText 3")

        builder = aw.DocumentBuilder(doc)

        # Create an AUTOTEXTLIST field and set the text that the field will display in Microsoft Word.
        # Set the text to prompt the user to right-click this field to select an AutoText building block,
        # whose contents the field will display.
        field = builder.insert_field(aw.fields.FieldType.FIELD_AUTO_TEXT_LIST, True).as_field_auto_text_list()
        field.entry_name = "Right click here to select an AutoText block"
        field.list_style = "Heading 1"
        field.screen_tip = "Hover tip text for AutoTextList goes here"

        self.assertEqual(" AUTOTEXTLIST  \"Right click here to select an AutoText block\" " +
                        "\\s \"Heading 1\" " +
                        "\\t \"Hover tip text for AutoTextList goes here\"", field.get_field_code())

        doc.save(ARTIFACTS_DIR + "Field.field_auto_text_list.dotx")
        self._test_field_auto_text_list(doc) #ExSkip

    @staticmethod
    def append_auto_text_entry(glossary_doc: aw.buildingblocks.GlossaryDocument, name: str, contents: str):
        """Create an AutoText-type building block and add it to a glossary document."""

        building_block = aw.buildingblocks.BuildingBlock(glossary_doc)
        building_block.name = name
        building_block.gallery = aw.buildingblocks.BuildingBlockGallery.AUTO_TEXT
        building_block.category = "General"
        building_block.behavior = aw.buildingblocks.BuildingBlockBehavior.PARAGRAPH

        section = aw.Section(glossary_doc)
        section.append_child(aw.Body(glossary_doc))
        section.body.append_paragraph(contents)
        building_block.append_child(section)

        glossary_doc.append_child(building_block)

    #ExEnd

    def _test_field_auto_text_list(self, doc: aw.Document):

        doc = DocumentHelper.save_open(doc)

        self.assertEqual(3, doc.glossary_document.count)
        self.assertEqual("AutoText 1", doc.glossary_document.building_blocks[0].name)
        self.assertEqual("Contents of AutoText 1", doc.glossary_document.building_blocks[0].get_text().strip())
        self.assertEqual("AutoText 2", doc.glossary_document.building_blocks[1].name)
        self.assertEqual("Contents of AutoText 2", doc.glossary_document.building_blocks[1].get_text().strip())
        self.assertEqual("AutoText 3", doc.glossary_document.building_blocks[2].name)
        self.assertEqual("Contents of AutoText 3", doc.glossary_document.building_blocks[2].get_text().strip())

        field = doc.range.fields[0].as_field_auto_text_list()

        self.verify_field(aw.fields.FieldType.FIELD_AUTO_TEXT_LIST,
            " AUTOTEXTLIST  \"Right click here to select an AutoText block\" \\s \"Heading 1\" \\t \"Hover tip text for AutoTextList goes here\"",
            "", field)
        self.assertEqual("Right click here to select an AutoText block", field.entry_name)
        self.assertEqual("Heading 1", field.list_style)
        self.assertEqual("Hover tip text for AutoTextList goes here", field.screen_tip)

    #def test_field_greeting_line(self):

    #    #ExStart
    #    #ExFor:FieldGreetingLine
    #    #ExFor:FieldGreetingLine.alternate_text
    #    #ExFor:FieldGreetingLine.get_field_names
    #    #ExFor:FieldGreetingLine.language_id
    #    #ExFor:FieldGreetingLine.name_format
    #    #ExSummary:Shows how to insert a GREETINGLINE field.
    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    # Create a generic greeting using a GREETINGLINE field, and some text after it.
    #    field = builder.insert_field(aw.fields.FieldType.FIELD_GREETING_LINE, True).as_field_greeting_line()
    #    builder.writeln("\n\n\tThis is your custom greeting, created programmatically using Aspose Words!")

    #    # A GREETINGLINE field accepts values from a data source during a mail merge, like a MERGEFIELD.
    #    # It can also format how the source's data is written in its place once the mail merge is complete.
    #    # The field names collection corresponds to the columns from the data source
    #    # that the field will take values from.
    #    self.assertEqual(0, len(field.get_field_names()))

    #    # To populate that array, we need to specify a format for our greeting line.
    #    field.name_format = "<< _BEFORE_ Dear >><< _TITLE0_ >><< _LAST0_ >><< _AFTER_ ,>> "

    #    # Now, our field will accept values from these two columns in the data source.
    #    self.assertEqual("Courtesy Title", field.get_field_names()[0])
    #    self.assertEqual("Last Name", field.get_field_names()[1])
    #    self.assertEqual(2, len(field.get_field_names()))

    #    # This string will cover any cases where the data table data is invalid
    #    # by substituting the malformed name with a string.
    #    field.alternate_text = "Sir or Madam"

    #    # Set a locale to format the result.
    #    field.language_id = "1033" # en-US

    #    self.assertEqual(" GREETINGLINE  \\f \"<< _BEFORE_ Dear >><< _TITLE0_ >><< _LAST0_ >><< _AFTER_ ,>> \" \\e \"Sir or Madam\" \\l 1033",
    #        field.get_field_code())

    #    # Create a data table with columns whose names match elements
    #    # from the field's field names collection, and then carry out the mail merge.
    #    table = DataTable("Employees")
    #    table.columns.add("Courtesy Title")
    #    table.columns.add("First Name")
    #    table.columns.add("Last Name")
    #    table.rows.add("Mr.", "John", "Doe")
    #    table.rows.add("Mrs.", "Jane", "Cardholder")

    #    # This row has an invalid value in the Courtesy Title column, so our greeting will default to the alternate text.
    #    table.rows.add("", "No", "Name")

    #    doc.mail_merge.execute(table)

    #    self.assertEqual(0, len(doc.range.fields))
    #    self.assertEqual("Dear Mr. Doe,\r\r\tThis is your custom greeting, created programmatically using Aspose Words!\r" +
    #                    "\fDear Mrs. Cardholder,\r\r\tThis is your custom greeting, created programmatically using Aspose Words!\r" +
    #                    "\fDear Sir or Madam,\r\r\tThis is your custom greeting, created programmatically using Aspose Words!",
    #        doc.get_text().strip())
    #    #ExEnd

    def test_field_list_num(self):

        #ExStart
        #ExFor:FieldListNum
        #ExFor:FieldListNum.has_list_name
        #ExFor:FieldListNum.list_level
        #ExFor:FieldListNum.list_name
        #ExFor:FieldListNum.starting_number
        #ExSummary:Shows how to number paragraphs with LISTNUM fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # LISTNUM fields display a number that increments at each LISTNUM field.
        # These fields also have a variety of options that allow us to use them to emulate numbered lists.
        field = builder.insert_field(aw.fields.FieldType.FIELD_LIST_NUM, True).as_field_list_num()

        # Lists start counting at 1 by default, but we can set this number to a different value, such as 0.
        # This field will display "0)".
        field.starting_number = "0"
        builder.writeln("Paragraph 1")

        self.assertEqual(" LISTNUM  \\s 0", field.get_field_code())

        # LISTNUM fields maintain separate counts for each list level.
        # Inserting a LISTNUM field in the same paragraph as another LISTNUM field
        # increases the list level instead of the count.
        # The next field will continue the count we started above and display a value of "1" at list level 1.
        builder.insert_field(aw.fields.FieldType.FIELD_LIST_NUM, True)

        # This field will start a count at list level 2. It will display a value of "1".
        builder.insert_field(aw.fields.FieldType.FIELD_LIST_NUM, True)

        # This field will start a count at list level 3. It will display a value of "1".
        # Different list levels have different formatting,
        # so these fields combined will display a value of "1)a)i)".
        builder.insert_field(aw.fields.FieldType.FIELD_LIST_NUM, True)
        builder.writeln("Paragraph 2")

        # The next LISTNUM field that we insert will continue the count at the list level
        # that the previous LISTNUM field was on.
        # We can use the "list_level" property to jump to a different list level.
        # If this LISTNUM field stayed on list level 3, it would display "ii)",
        # but, since we have moved it to list level 2, it carries on the count at that level and displays "b)".
        field = builder.insert_field(aw.fields.FieldType.FIELD_LIST_NUM, True).as_field_list_num()
        field.list_level = "2"
        builder.writeln("Paragraph 3")

        self.assertEqual(" LISTNUM  \\l 2", field.get_field_code())

        # We can set the list_name property to get the field to emulate a different AUTONUM field type.
        # "NumberDefault" emulates AUTONUM, "OutlineDefault" emulates AUTONUMOUT,
        # and "LegalDefault" emulates AUTONUMLGL fields.
        # The "OutlineDefault" list name with 1 as the starting number will result in displaying "I.".
        field = builder.insert_field(aw.fields.FieldType.FIELD_LIST_NUM, True).as_field_list_num()
        field.starting_number = "1"
        field.list_name = "OutlineDefault"
        builder.writeln("Paragraph 4")

        self.assertTrue(field.has_list_name)
        self.assertEqual(" LISTNUM  OutlineDefault \\s 1", field.get_field_code())

        # The list_name does not carry over from the previous field, so we will need to set it for each new field.
        # This field continues the count with the different list name and displays "II.".
        field = builder.insert_field(aw.fields.FieldType.FIELD_LIST_NUM, True).as_field_list_num()
        field.list_name = "OutlineDefault"
        builder.writeln("Paragraph 5")

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_list_num.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_list_num.docx")

        self.assertEqual(7, doc.range.fields.count)

        field = doc.range.fields[0].as_field_list_num()

        self.verify_field(aw.fields.FieldType.FIELD_LIST_NUM, " LISTNUM  \\s 0", "", field)
        self.assertEqual("0", field.starting_number)
        self.assertIsNone(field.list_level)
        self.assertFalse(field.has_list_name)
        self.assertIsNone(field.list_name)

        for i in range(1, 4):

            field = doc.range.fields[i].as_field_list_num()

            self.verify_field(aw.fields.FieldType.FIELD_LIST_NUM, " LISTNUM ", "", field)
            self.assertIsNone(field.starting_number)
            self.assertIsNone(field.list_level)
            self.assertFalse(field.has_list_name)
            self.assertIsNone(field.list_name)

        field = doc.range.fields[4].as_field_list_num()

        self.verify_field(aw.fields.FieldType.FIELD_LIST_NUM, " LISTNUM  \\l 2", "", field)
        self.assertIsNone(field.starting_number)
        self.assertEqual("2", field.list_level)
        self.assertFalse(field.has_list_name)
        self.assertIsNone(field.list_name)

        field = doc.range.fields[5].as_field_list_num()

        self.verify_field(aw.fields.FieldType.FIELD_LIST_NUM, " LISTNUM  OutlineDefault \\s 1", "", field)
        self.assertEqual("1", field.starting_number)
        self.assertIsNone(field.list_level)
        self.assertTrue(field.has_list_name)
        self.assertEqual("OutlineDefault", field.list_name)

    #def test_merge_field(self):

    #    #ExStart
    #    #ExFor:FieldMergeField
    #    #ExFor:FieldMergeField.field_name
    #    #ExFor:FieldMergeField.field_name_no_prefix
    #    #ExFor:FieldMergeField.is_mapped
    #    #ExFor:FieldMergeField.is_vertical_formatting
    #    #ExFor:FieldMergeField.text_after
    #    #ExFor:FieldMergeField.text_before
    #    #ExSummary:Shows how to use MERGEFIELD fields to perform a mail merge.
    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    # Create a data table to be used as a mail merge data source.
    #    table = DataTable("Employees")
    #    table.columns.add("Courtesy Title")
    #    table.columns.add("First Name")
    #    table.columns.add("Last Name")
    #    table.rows.add("Mr.", "John", "Doe")
    #    table.rows.add("Mrs.", "Jane", "Cardholder")

    #    # Insert a MERGEFIELD with a FieldName property set to the name of a column in the data source.
    #    field_merge_field = builder.insert_field(aw.fields.FieldType.FIELD_MERGE_FIELD, True).as_field_merge_field()
    #    field_merge_field.field_name = "Courtesy Title"
    #    field_merge_field.is_mapped = True
    #    field_merge_field.is_vertical_formatting = False

    #    # We can apply text before and after the value that this field accepts when the merge takes place.
    #    field_merge_field.text_before = "Dear "
    #    field_merge_field.text_after = " "

    #    self.assertEqual(" MERGEFIELD  \"Courtesy Title\" \\m \\b \"Dear \" \\f \" \"", field_merge_field.get_field_code())

    #    # Insert another MERGEFIELD for a different column in the data source.
    #    field_merge_field = builder.insert_field(aw.fields.FieldType.FIELD_MERGE_FIELD, True).as_field_merge_field()
    #    field_merge_field.field_name = "Last Name"
    #    field_merge_field.text_after = ":"

    #    doc.update_fields()
    #    doc.mail_merge.execute(table)

    #    self.assertEqual("Dear Mr. Doe:\u000cDear Mrs. Cardholder:", doc.get_text().strip())
    #    #ExEnd

    #    self.assertEqual(0, len(doc.range.fields))

    #ExStart
    #ExFor:FieldToc
    #ExFor:FieldToc.bookmark_name
    #ExFor:FieldToc.custom_styles
    #ExFor:FieldToc.entry_separator
    #ExFor:FieldToc.heading_level_range
    #ExFor:FieldToc.hide_in_web_layout
    #ExFor:FieldToc.insert_hyperlinks
    #ExFor:FieldToc.page_number_omitting_level_range
    #ExFor:FieldToc.preserve_line_breaks
    #ExFor:FieldToc.preserve_tabs
    #ExFor:FieldToc.update_page_numbers
    #ExFor:FieldToc.use_paragraph_outline_level
    #ExFor:FieldOptions.custom_toc_style_separator
    #ExSummary:Shows how to insert a TOC, and populate it with entries based on heading styles.
    def test_field_toc(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_bookmark("MyBookmark")

        # Insert a TOC field, which will compile all headings into a table of contents.
        # For each heading, this field will create a line with the text in that heading style to the left,
        # and the page the heading appears on to the right.
        field = builder.insert_field(aw.fields.FieldType.FIELD_TOC, True).as_field_toc()

        # Use the "bookmark_name" property to only list headings
        # that appear within the bounds of a bookmark with the "MyBookmark" name.
        field.bookmark_name = "MyBookmark"

        # Text with a built-in heading style, such as "Heading 1", applied to it will count as a heading.
        # We can name additional styles to be picked up as headings by the TOC in this property and their TOC levels.
        field.custom_styles = "Quote; 6; Intense Quote; 7"

        # By default, Styles/TOC levels are separated in the "custom_styles" property by a comma,
        # but we can set a custom delimiter in this property.
        doc.field_options.custom_toc_style_separator = ";"

        # Configure the field to exclude any headings that have TOC levels outside of this range.
        field.heading_level_range = "1-3"

        # The TOC will not display the page numbers of headings whose TOC levels are within this range.
        field.page_number_omitting_level_range = "2-5"

        # Set a custom string that will separate every heading from its page number.
        field.entry_separator = "-"
        field.insert_hyperlinks = True
        field.hide_in_web_layout = False
        field.preserve_line_breaks = True
        field.preserve_tabs = True
        field.use_paragraph_outline_level = False

        ExField.insert_new_page_with_heading(builder, "First entry", "Heading 1")
        builder.writeln("Paragraph text.")
        ExField.insert_new_page_with_heading(builder, "Second entry", "Heading 1")
        ExField.insert_new_page_with_heading(builder, "Third entry", "Quote")
        ExField.insert_new_page_with_heading(builder, "Fourth entry", "Intense Quote")

        # These two headings will have the page numbers omitted because they are within the "2-5" range.
        ExField.insert_new_page_with_heading(builder, "Fifth entry", "Heading 2")
        ExField.insert_new_page_with_heading(builder, "Sixth entry", "Heading 3")

        # This entry does not appear because "Heading 4" is outside of the "1-3" range that we have set earlier.
        ExField.insert_new_page_with_heading(builder, "Seventh entry", "Heading 4")

        builder.end_bookmark("MyBookmark")
        builder.writeln("Paragraph text.")

        # This entry does not appear because it is outside the bookmark specified by the TOC.
        ExField.insert_new_page_with_heading(builder, "Eighth entry", "Heading 1")

        self.assertEqual(" TOC  \\b MyBookmark \\t \"Quote; 6; Intense Quote; 7\" \\o 1-3 \\n 2-5 \\p - \\h \\x \\w", field.get_field_code())

        field.update_page_numbers()
        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_toc.docx")
        self._test_field_toc(doc) #ExSkip

    @staticmethod
    def insert_new_page_with_heading(builder: aw.DocumentBuilder, caption_text: str, style_name: str):
        """Start a new page and insert a paragraph of a specified style."""

        builder.insert_break(aw.BreakType.PAGE_BREAK)
        original_style = builder.paragraph_format.style_name
        builder.paragraph_format.style = builder.document.styles.get_by_name(style_name)
        builder.writeln(caption_text)
        builder.paragraph_format.style = builder.document.styles.get_by_name(original_style)

    #ExEnd

    def _test_field_toc(self, doc: aw.Document):

        doc = DocumentHelper.save_open(doc)
        field = doc.range.fields[0].as_field_toc()

        self.assertEqual("MyBookmark", field.bookmark_name)
        self.assertEqual("Quote; 6; Intense Quote; 7", field.custom_styles)
        self.assertEqual("-", field.entry_separator)
        self.assertEqual("1-3", field.heading_level_range)
        self.assertEqual("2-5", field.page_number_omitting_level_range)
        self.assertFalse(field.hide_in_web_layout)
        self.assertTrue(field.insert_hyperlinks)
        self.assertTrue(field.preserve_line_breaks)
        self.assertTrue(field.preserve_tabs)
        self.assertTrue(field.update_page_numbers())
        self.assertFalse(field.use_paragraph_outline_level)
        self.assertEqual(" TOC  \\b MyBookmark \\t \"Quote; 6; Intense Quote; 7\" \\o 1-3 \\n 2-5 \\p - \\h \\x \\w", field.get_field_code())
        self.assertEqual("\u0013 HYPERLINK \\l \"_Toc256000001\" \u0014First entry-\u0013 PAGEREF _Toc256000001 \\h \u00142\u0015\u0015\r" +
                        "\u0013 HYPERLINK \\l \"_Toc256000002\" \u0014Second entry-\u0013 PAGEREF _Toc256000002 \\h \u00143\u0015\u0015\r" +
                        "\u0013 HYPERLINK \\l \"_Toc256000003\" \u0014Third entry-\u0013 PAGEREF _Toc256000003 \\h \u00144\u0015\u0015\r" +
                        "\u0013 HYPERLINK \\l \"_Toc256000004\" \u0014Fourth entry-\u0013 PAGEREF _Toc256000004 \\h \u00145\u0015\u0015\r" +
                        "\u0013 HYPERLINK \\l \"_Toc256000005\" \u0014Fifth entry\u0015\r" +
                        "\u0013 HYPERLINK \\l \"_Toc256000006\" \u0014Sixth entry\u0015\r", field.result)

    #ExStart
    #ExFor:FieldToc.entry_identifier
    #ExFor:FieldToc.entry_level_range
    #ExFor:FieldTC
    #ExFor:FieldTC.omit_page_number
    #ExFor:FieldTC.text
    #ExFor:FieldTC.type_identifier
    #ExFor:FieldTC.entry_level
    #ExSummary:Shows how to insert a TOC field, and filter which TC fields end up as entries.
    def test_field_toc_entry_identifier(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a TOC field, which will compile all TC fields into a table of contents.
        field_toc = builder.insert_field(aw.fields.FieldType.FIELD_TOC, True).as_field_toc()

        # Configure the field only to pick up TC entries of the "A" type, and an entry-level between 1 and 3.
        field_toc.entry_identifier = "A"
        field_toc.entry_level_range = "1-3"

        self.assertEqual(" TOC  \\f A \\l 1-3", field_toc.get_field_code())

        # These two entries will appear in the table.
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        ExField.insert_toc_entry(builder, "TC field 1", "A", "1")
        ExField.insert_toc_entry(builder, "TC field 2", "A", "2")

        self.assertEqual(" TC  \"TC field 1\" \\n \\f A \\l 1", doc.range.fields[1].get_field_code())

        # This entry will be omitted from the table because it has a different type from "A".
        ExField.insert_toc_entry(builder, "TC field 3", "B", "1")

        # This entry will be omitted from the table because it has an entry-level outside of the 1-3 range.
        ExField.insert_toc_entry(builder, "TC field 4", "A", "5")

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.tc.docx")
        self._test_field_toc_entry_identifier(doc) #ExSkip

    @staticmethod
    def insert_toc_entry(builder: aw.DocumentBuilder, text: str, type_identifier: str, entry_level: str):
        """Use a document builder to insert a TC field."""

        field_tc = builder.insert_field(aw.fields.FieldType.FIELD_TOCENTRY, True).as_field_tc()
        field_tc.omit_page_number = True
        field_tc.text = text
        field_tc.type_identifier = type_identifier
        field_tc.entry_level = entry_level

    #ExEnd

    def _test_field_toc_entry_identifier(self, doc: aw.Document):

        doc = DocumentHelper.save_open(doc)
        field_toc = doc.range.fields[0].as_field_toc()

        self.verify_field(aw.fields.FieldType.FIELD_TOC, " TOC  \\f A \\l 1-3", "TC field 1\rTC field 2\r", field_toc)
        self.assertEqual("A", field_toc.entry_identifier)
        self.assertEqual("1-3", field_toc.entry_level_range)

        field_tc = doc.range.fields[1].as_field_tc()

        self.verify_field(aw.fields.FieldType.FIELD_TOCENTRY, " TC  \"TC field 1\" \\n \\f A \\l 1", "", field_tc)
        self.assertTrue(field_tc.omit_page_number)
        self.assertEqual("TC field 1", field_tc.text)
        self.assertEqual("A", field_tc.type_identifier)
        self.assertEqual("1", field_tc.entry_level)

        field_tc = doc.range.fields[2].as_field_tc()

        self.verify_field(aw.fields.FieldType.FIELD_TOCENTRY, " TC  \"TC field 2\" \\n \\f A \\l 2", "", field_tc)
        self.assertTrue(field_tc.omit_page_number)
        self.assertEqual("TC field 2", field_tc.text)
        self.assertEqual("A", field_tc.type_identifier)
        self.assertEqual("2", field_tc.entry_level)

        field_tc = doc.range.fields[3].as_field_tc()

        self.verify_field(aw.fields.FieldType.FIELD_TOCENTRY, " TC  \"TC field 3\" \\n \\f B \\l 1", "", field_tc)
        self.assertTrue(field_tc.omit_page_number)
        self.assertEqual("TC field 3", field_tc.text)
        self.assertEqual("B", field_tc.type_identifier)
        self.assertEqual("1", field_tc.entry_level)

        field_tc = doc.range.fields[4].as_field_tc()

        self.verify_field(aw.fields.FieldType.FIELD_TOCENTRY, " TC  \"TC field 4\" \\n \\f A \\l 5", "", field_tc)
        self.assertTrue(field_tc.omit_page_number)
        self.assertEqual("TC field 4", field_tc.text)
        self.assertEqual("A", field_tc.type_identifier)
        self.assertEqual("5", field_tc.entry_level)

    def test_toc_seq_prefix(self):

        #ExStart
        #ExFor:FieldToc
        #ExFor:FieldToc.table_of_figures_label
        #ExFor:FieldToc.prefixed_sequence_identifier
        #ExFor:FieldToc.sequence_separator
        #ExFor:FieldSeq
        #ExFor:FieldSeq.sequence_identifier
        #ExSummary:Shows how to populate a TOC field with entries using SEQ fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # A TOC field can create an entry in its table of contents for each SEQ field found in the document.
        # Each entry contains the paragraph that includes the SEQ field and the page's number that the field appears on.
        field_toc = builder.insert_field(aw.fields.FieldType.FIELD_TOC, True).as_field_toc()

        # SEQ fields display a count that increments at each SEQ field.
        # These fields also maintain separate counts for each unique named sequence
        # identified by the SEQ field's "sequence_identifier" property.
        # Use the "table_of_figures_label" property to name a main sequence for the TOC.
        # Now, this TOC will only create entries out of SEQ fields with their "sequence_identifier" set to "MySequence".
        field_toc.table_of_figures_label = "MySequence"

        # We can name another SEQ field sequence in the "prefixed_sequence_identifier" property.
        # SEQ fields from this prefix sequence will not create TOC entries.
        # Every TOC entry created from a main sequence SEQ field will now also display the count that
        # the prefix sequence is currently on at the primary sequence SEQ field that made the entry.
        field_toc.prefixed_sequence_identifier = "PrefixSequence"

        # Each TOC entry will display the prefix sequence count immediately to the left
        # of the page number that the main sequence SEQ field appears on.
        # We can specify a custom separator that will appear between these two numbers.
        field_toc.sequence_separator = ">"

        self.assertEqual(" TOC  \\c MySequence \\s PrefixSequence \\d >", field_toc.get_field_code())

        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # There are two ways of using SEQ fields to populate this TOC.
        # 1 -  Inserting a SEQ field that belongs to the TOC's prefix sequence:
        # This field will increment the SEQ sequence count for the "PrefixSequence" by 1.
        # Since this field does not belong to the main sequence identified
        # by the "table_of_figures_label" property of the TOC, it will not appear as an entry.
        field_seq = builder.insert_field(aw.fields.FieldType.FIELD_SEQUENCE, True).as_field_seq()
        field_seq.sequence_identifier = "PrefixSequence"
        builder.insert_paragraph()

        self.assertEqual(" SEQ  PrefixSequence", field_seq.get_field_code())

        # 2 -  Inserting a SEQ field that belongs to the TOC's main sequence:
        # This SEQ field will create an entry in the TOC.
        # The TOC entry will contain the paragraph that the SEQ field is in and the number of the page that it appears on.
        # This entry will also display the count that the prefix sequence is currently at,
        # separated from the page number by the value in the TOC's "seqence_separator" property.
        # The "PrefixSequence" count is at 1, this main sequence SEQ field is on page 2,
        # and the separator is ">", so entry will display "1>2".
        builder.write("First TOC entry, MySequence #")
        field_seq = builder.insert_field(aw.fields.FieldType.FIELD_SEQUENCE, True).as_field_seq()
        field_seq.sequence_identifier = "MySequence"

        self.assertEqual(" SEQ  MySequence", field_seq.get_field_code())

        # Insert a page, advance the prefix sequence by 2, and insert a SEQ field to create a TOC entry afterwards.
        # The prefix sequence is now at 2, and the main sequence SEQ field is on page 3,
        # so the TOC entry will display "2>3" at its page count.
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        field_seq = builder.insert_field(aw.fields.FieldType.FIELD_SEQUENCE, True).as_field_seq()
        field_seq.sequence_identifier = "PrefixSequence"
        builder.insert_paragraph()
        field_seq = builder.insert_field(aw.fields.FieldType.FIELD_SEQUENCE, True).as_field_seq()
        builder.write("Second TOC entry, MySequence #")
        field_seq.sequence_identifier = "MySequence"

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.toc_seq_prefix.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.toc_seq_prefix.docx")

        self.assertEqual(9, doc.range.fields.count)

        field_toc = doc.range.fields[0].as_field_toc()
        print(field_toc.display_result)
        self.verify_field(aw.fields.FieldType.FIELD_TOC, " TOC  \\c MySequence \\s PrefixSequence \\d >",
            "First TOC entry, MySequence #12\t\u0013 SEQ PrefixSequence _Toc256000000 \\* ARABIC \u00141\u0015>\u0013 PAGEREF _Toc256000000 \\h \u00142\u0015\r2" +
            "Second TOC entry, MySequence #\t\u0013 SEQ PrefixSequence _Toc256000001 \\* ARABIC \u00142\u0015>\u0013 PAGEREF _Toc256000001 \\h \u00143\u0015\r",
            field_toc)
        self.assertEqual("MySequence", field_toc.table_of_figures_label)
        self.assertEqual("PrefixSequence", field_toc.prefixed_sequence_identifier)
        self.assertEqual(">", field_toc.sequence_separator)

        field_seq = doc.range.fields[1].as_field_seq()

        self.verify_field(aw.fields.FieldType.FIELD_SEQUENCE, " SEQ PrefixSequence _Toc256000000 \\* ARABIC ", "1", field_seq)
        self.assertEqual("PrefixSequence", field_seq.sequence_identifier)

        # Byproduct field created by Aspose.Words
        field_page_ref = doc.range.fields[2].as_field_page_ref()

        self.verify_field(aw.fields.FieldType.FIELD_PAGE_REF, " PAGEREF _Toc256000000 \\h ", "2", field_page_ref)
        self.assertEqual("PrefixSequence", field_seq.sequence_identifier)
        self.assertEqual("_Toc256000000", field_page_ref.bookmark_name)

        field_seq = doc.range.fields[3].as_field_seq()

        self.verify_field(aw.fields.FieldType.FIELD_SEQUENCE, " SEQ PrefixSequence _Toc256000001 \\* ARABIC ", "2", field_seq)
        self.assertEqual("PrefixSequence", field_seq.sequence_identifier)

        field_page_ref = doc.range.fields[4].as_field_page_ref()

        self.verify_field(aw.fields.FieldType.FIELD_PAGE_REF, " PAGEREF _Toc256000001 \\h ", "3", field_page_ref)
        self.assertEqual("PrefixSequence", field_seq.sequence_identifier)
        self.assertEqual("_Toc256000001", field_page_ref.bookmark_name)

        field_seq = doc.range.fields[5].as_field_seq()

        self.verify_field(aw.fields.FieldType.FIELD_SEQUENCE, " SEQ  PrefixSequence", "1", field_seq)
        self.assertEqual("PrefixSequence", field_seq.sequence_identifier)

        field_seq = doc.range.fields[6].as_field_seq()

        self.verify_field(aw.fields.FieldType.FIELD_SEQUENCE, " SEQ  MySequence", "1", field_seq)
        self.assertEqual("MySequence", field_seq.sequence_identifier)

        field_seq = doc.range.fields[7].as_field_seq()

        self.verify_field(aw.fields.FieldType.FIELD_SEQUENCE, " SEQ  PrefixSequence", "2", field_seq)
        self.assertEqual("PrefixSequence", field_seq.sequence_identifier)

        field_seq = doc.range.fields[8].as_field_seq()

        self.verify_field(aw.fields.FieldType.FIELD_SEQUENCE, " SEQ  MySequence", "2", field_seq)
        self.assertEqual("MySequence", field_seq.sequence_identifier)

    def test_toc_seq_numbering(self):

        #ExStart
        #ExFor:FieldSeq
        #ExFor:FieldSeq.insert_next_number
        #ExFor:FieldSeq.reset_heading_level
        #ExFor:FieldSeq.reset_number
        #ExFor:FieldSeq.sequence_identifier
        #ExSummary:Shows create numbering using SEQ fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # SEQ fields display a count that increments at each SEQ field.
        # These fields also maintain separate counts for each unique named sequence
        # identified by the SEQ field's "sequence_identifier" property.
        # Insert a SEQ field that will display the current count value of "MySequence",
        # after using the "reset_number" property to set it to 100.
        builder.write("#")
        field_seq = builder.insert_field(aw.fields.FieldType.FIELD_SEQUENCE, True).as_field_seq()
        field_seq.sequence_identifier = "MySequence"
        field_seq.reset_number = "100"
        field_seq.update()

        self.assertEqual(" SEQ  MySequence \\r 100", field_seq.get_field_code())
        self.assertEqual("100", field_seq.result)

        # Display the next number in this sequence with another SEQ field.
        builder.write(", #")
        field_seq = builder.insert_field(aw.fields.FieldType.FIELD_SEQUENCE, True).as_field_seq()
        field_seq.sequence_identifier = "MySequence"
        field_seq.update()

        self.assertEqual("101", field_seq.result)

        # Insert a level 1 heading.
        builder.insert_break(aw.BreakType.PARAGRAPH_BREAK)
        builder.paragraph_format.style = doc.styles.get_by_name("Heading 1")
        builder.writeln("This level 1 heading will reset MySequence to 1")
        builder.paragraph_format.style = doc.styles.get_by_name("Normal")

        # Insert another SEQ field from the same sequence and configure it to reset the count at every heading with 1.
        builder.write("\n#")
        field_seq = builder.insert_field(aw.fields.FieldType.FIELD_SEQUENCE, True).as_field_seq()
        field_seq.sequence_identifier = "MySequence"
        field_seq.reset_heading_level = "1"
        field_seq.update()

        # The above heading is a level 1 heading, so the count for this sequence is reset to 1.
        self.assertEqual(" SEQ  MySequence \\s 1", field_seq.get_field_code())
        self.assertEqual("1", field_seq.result)

        # Move to the next number of this sequence.
        builder.write(", #")
        field_seq = builder.insert_field(aw.fields.FieldType.FIELD_SEQUENCE, True).as_field_seq()
        field_seq.sequence_identifier = "MySequence"
        field_seq.insert_next_number = True
        field_seq.update()

        self.assertEqual(" SEQ  MySequence \\n", field_seq.get_field_code())
        self.assertEqual("2", field_seq.result)

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.toc_seq_numbering.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.toc_seq_numbering.docx")

        self.assertEqual(4, doc.range.fields.count)

        field_seq = doc.range.fields[0].as_field_seq()

        self.verify_field(aw.fields.FieldType.FIELD_SEQUENCE, " SEQ  MySequence \\r 100", "100", field_seq)
        self.assertEqual("MySequence", field_seq.sequence_identifier)

        field_seq = doc.range.fields[1].as_field_seq()

        self.verify_field(aw.fields.FieldType.FIELD_SEQUENCE, " SEQ  MySequence", "101", field_seq)
        self.assertEqual("MySequence", field_seq.sequence_identifier)

        field_seq = doc.range.fields[2].as_field_seq()

        self.verify_field(aw.fields.FieldType.FIELD_SEQUENCE, " SEQ  MySequence \\s 1", "1", field_seq)
        self.assertEqual("MySequence", field_seq.sequence_identifier)

        field_seq = doc.range.fields[3].as_field_seq()

        self.verify_field(aw.fields.FieldType.FIELD_SEQUENCE, " SEQ  MySequence \\n", "2", field_seq)
        self.assertEqual("MySequence", field_seq.sequence_identifier)

    @unittest.skip("WORDSNET-18083")
    def test_toc_seq_bookmark(self):

        #ExStart
        #ExFor:FieldSeq
        #ExFor:FieldSeq.bookmark_name
        #ExSummary:Shows how to combine table of contents and sequence fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # A TOC field can create an entry in its table of contents for each SEQ field found in the document.
        # Each entry contains the paragraph that contains the SEQ field,
        # and the number of the page that the field appears on.
        field_toc = builder.insert_field(aw.fields.FieldType.FIELD_TOC, True).as_field_toc()

        # Configure this TOC field to have a "sequence_identifier" property with a value of "MySequence".
        field_toc.table_of_figures_label = "MySequence"

        # Configure this TOC field to only pick up SEQ fields that are within the bounds of a bookmark
        # named "TOCBookmark".
        field_toc.bookmark_name = "TOCBookmark"
        builder.insert_break(aw.BreakType.PAGE_BREAK)

        self.assertEqual(" TOC  \\c MySequence \\b TOCBookmark", field_toc.get_field_code())

        # SEQ fields display a count that increments at each SEQ field.
        # These fields also maintain separate counts for each unique named sequence
        # identified by the SEQ field's "sequence_identifier" property.
        # Insert a SEQ field that has a sequence identifier that matches the TOC's
        # "table_of_figures_label" property. This field will not create an entry in the TOC since it is outside
        # the bookmark's bounds designated by "BookmarkName".
        builder.write("MySequence #")
        field_seq = builder.insert_field(aw.fields.FieldType.FIELD_SEQUENCE, True).as_field_seq()
        field_seq.sequence_identifier = "MySequence"
        builder.writeln(", will not show up in the TOC because it is outside of the bookmark.")

        builder.start_bookmark("TOCBookmark")

        # This SEQ field's sequence matches the TOC's "table_of_figures_label" property and is within the bookmark's bounds.
        # The paragraph that contains this field will show up in the TOC as an entry.
        builder.write("MySequence #")
        field_seq = builder.insert_field(aw.fields.FieldType.FIELD_SEQUENCE, True).as_field_seq()
        field_seq.sequence_identifier = "MySequence"
        builder.writeln(", will show up in the TOC next to the entry for the above caption.")

        # This SEQ field's sequence does not match the TOC's "table_of_figures_label" property,
        # and is within the bounds of the bookmark. Its paragraph will not show up in the TOC as an entry.
        builder.write("MySequence #")
        field_seq = builder.insert_field(aw.fields.FieldType.FIELD_SEQUENCE, True).as_field_seq()
        field_seq.sequence_identifier = "OtherSequence"
        builder.writeln(", will not show up in the TOC because it's from a different sequence identifier.")

        # This SEQ field's sequence matches the TOC's "table_of_figures_label" property and is within the bounds of the bookmark.
        # This field also references another bookmark. The contents of that bookmark will appear in the TOC entry for this SEQ field.
        # The SEQ field itself will not display the contents of that bookmark.
        field_seq = builder.insert_field(aw.fields.FieldType.FIELD_SEQUENCE, True).as_field_seq()
        field_seq.sequence_identifier = "MySequence"
        field_seq.bookmark_name = "SEQBookmark"
        self.assertEqual(" SEQ  MySequence SEQBookmark", field_seq.get_field_code())

        # Create a bookmark with contents that will show up in the TOC entry due to the above SEQ field referencing it.
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.start_bookmark("SEQBookmark")
        builder.write("MySequence #")
        field_seq = builder.insert_field(aw.fields.FieldType.FIELD_SEQUENCE, True).as_field_seq()
        field_seq.sequence_identifier = "MySequence"
        builder.writeln(", text from inside SEQBookmark.")
        builder.end_bookmark("SEQBookmark")

        builder.end_bookmark("TOCBookmark")

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.toc_seq_bookmark.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.toc_seq_bookmark.docx")

        self.assertEqual(8, doc.range.fields.count)

        field_toc = doc.range.fields[0].as_field_toc()
        page_ref_ids = [s for s in field_toc.result.split(' ') if s.startswith("_Toc")]

        self.assertEqual(aw.fields.FieldType.FIELD_TOC, field_toc.type)
        self.assertEqual("MySequence", field_toc.table_of_figures_label)
        self.verify_field(aw.fields.FieldType.FIELD_TOC, " TOC  \\c MySequence \\b TOCBookmark",
            f"MySequence #2, will show up in the TOC next to the entry for the above caption.\t\u0013 PAGEREF {page_ref_ids[0]} \\h \u00142\u0015\r" +
            f"3MySequence #3, text from inside SEQBookmark.\t\u0013 PAGEREF {page_ref_ids[1]} \\h \u00142\u0015\r", field_toc)

        field_page_ref = doc.range.fields[1].as_field_page_ref()

        self.verify_field(aw.fields.FieldType.FIELD_PAGE_REF, f" PAGEREF {page_ref_ids[0]} \\h ", "2", field_page_ref)
        self.assertEqual(page_ref_ids[0], field_page_ref.bookmark_name)

        field_page_ref = doc.range.fields[2].as_field_page_ref()

        self.verify_field(aw.fields.FieldType.FIELD_PAGE_REF, f" PAGEREF {page_ref_ids[1]} \\h ", "2", field_page_ref)
        self.assertEqual(page_ref_ids[1], field_page_ref.bookmark_name)

        field_seq = doc.range.fields[3].as_field_seq()

        self.verify_field(aw.fields.FieldType.FIELD_SEQUENCE, " SEQ  MySequence", "1", field_seq)
        self.assertEqual("MySequence", field_seq.sequence_identifier)

        field_seq = doc.range.fields[4].as_field_seq()

        self.verify_field(aw.fields.FieldType.FIELD_SEQUENCE, " SEQ  MySequence", "2", field_seq)
        self.assertEqual("MySequence", field_seq.sequence_identifier)

        field_seq = doc.range.fields[5].as_field_seq()

        self.verify_field(aw.fields.FieldType.FIELD_SEQUENCE, " SEQ  OtherSequence", "1", field_seq)
        self.assertEqual("OtherSequence", field_seq.sequence_identifier)

        field_seq = doc.range.fields[6].as_field_seq()

        self.verify_field(aw.fields.FieldType.FIELD_SEQUENCE, " SEQ  MySequence SEQBookmark", "3", field_seq)
        self.assertEqual("MySequence", field_seq.sequence_identifier)
        self.assertEqual("SEQBookmark", field_seq.bookmark_name)

        field_seq = doc.range.fields[7].as_field_seq()

        self.verify_field(aw.fields.FieldType.FIELD_SEQUENCE, " SEQ  MySequence", "3", field_seq)
        self.assertEqual("MySequence", field_seq.sequence_identifier)

    @unittest.skip("WORDSNET-13854")
    def test_field_citation(self):

        #ExStart
        #ExFor:FieldCitation
        #ExFor:FieldCitation.another_source_tag
        #ExFor:FieldCitation.format_language_id
        #ExFor:FieldCitation.page_number
        #ExFor:FieldCitation.prefix
        #ExFor:FieldCitation.source_tag
        #ExFor:FieldCitation.suffix
        #ExFor:FieldCitation.suppress_author
        #ExFor:FieldCitation.suppress_title
        #ExFor:FieldCitation.suppress_year
        #ExFor:FieldCitation.volume_number
        #ExFor:FieldBibliography
        #ExFor:FieldBibliography.format_language_id
        #ExSummary:Shows how to work with CITATION and BIBLIOGRAPHY fields.
        # Open a document containing bibliographical sources that we can find in
        # Microsoft Word via References -> Citations & Bibliography -> Manage Sources.
        doc = aw.Document(MY_DIR + "Bibliography.docx")
        self.assertEqual(2, doc.range.fields.count) #ExSkip

        builder = aw.DocumentBuilder(doc)
        builder.write("Text to be cited with one source.")

        # Create a citation with just the page number and the author of the referenced book.
        field_citation = builder.insert_field(aw.fields.FieldType.FIELD_CITATION, True).as_field_citation()

        # We refer to sources using their tag names.
        field_citation.source_tag = "Book1"
        field_citation.page_number = "85"
        field_citation.suppress_author = False
        field_citation.suppress_title = True
        field_citation.suppress_year = True

        self.assertEqual(" CITATION  Book1 \\p 85 \\t \\y", field_citation.get_field_code())

        # Create a more detailed citation which cites two sources.
        builder.insert_paragraph()
        builder.write("Text to be cited with two sources.")
        field_citation = builder.insert_field(aw.fields.FieldType.FIELD_CITATION, True).as_field_citation()
        field_citation.source_tag = "Book1"
        field_citation.another_source_tag = "Book2"
        field_citation.format_language_id = "en-US"
        field_citation.page_number = "19"
        field_citation.prefix = "Prefix "
        field_citation.suffix = " Suffix"
        field_citation.suppress_author = False
        field_citation.suppress_title = False
        field_citation.suppress_year = False
        field_citation.volume_number = "VII"

        self.assertEqual(" CITATION  Book1 \\m Book2 \\l en-US \\p 19 \\f \"Prefix \" \\s \" Suffix\" \\v VII", field_citation.get_field_code())

        # We can use a BIBLIOGRAPHY field to display all the sources within the document.
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        field_bibliography = builder.insert_field(aw.fields.FieldType.FIELD_BIBLIOGRAPHY, True).as_field_bibliography()
        field_bibliography.format_language_id = "1124"

        self.assertEqual(" BIBLIOGRAPHY  \\l 1124", field_bibliography.get_field_code())

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_citation.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_citation.docx")

        self.assertEqual(5, doc.range.fields.count)

        field_citation = doc.range.fields[0].as_field_citation()

        self.verify_field(aw.fields.FieldType.FIELD_CITATION, " CITATION  Book1 \\p 85 \\t \\y", " (Doe, p. 85)", field_citation)
        self.assertEqual("Book1", field_citation.source_tag)
        self.assertEqual("85", field_citation.page_number)
        self.assertFalse(field_citation.suppress_author)
        self.assertTrue(field_citation.suppress_title)
        self.assertTrue(field_citation.suppress_year)

        field_citation = doc.range.fields[1].as_field_Citation()

        self.verify_field(aw.fields.FieldType.FIELD_CITATION,
            " CITATION  Book1 \\m Book2 \\l en-US \\p 19 \\f \"Prefix \" \\s \" Suffix\" \\v VII",
            " (Doe, 2018; Prefix Cardholder, 2018, VII:19 Suffix)", field_citation)
        self.assertEqual("Book1", field_citation.source_tag)
        self.assertEqual("Book2", field_citation.another_source_tag)
        self.assertEqual("en-US", field_citation.format_language_id)
        self.assertEqual("Prefix ", field_citation.prefix)
        self.assertEqual(" Suffix", field_citation.suffix)
        self.assertEqual("19", field_citation.page_number)
        self.assertFalse(field_citation.suppress_author)
        self.assertFalse(field_citation.suppress_title)
        self.assertFalse(field_citation.suppress_year)
        self.assertEqual("VII", field_citation.volume_number)

        field_bibliography = doc.range.fields[2].as_field_bibliography()

        self.verify_field(aw.fields.FieldType.FIELD_BIBLIOGRAPHY, " BIBLIOGRAPHY  \\l 1124",
            "Cardholder, A. (2018). My Book, Vol. II. New York: Doe Co. Ltd.\rDoe, J. (2018). My Book, Vol I. London: Doe Co. Ltd.\r", field_bibliography)
        self.assertEqual("1124", field_bibliography.format_language_id)

        field_citation = doc.range.fields[3].as_field_citation()

        self.verify_field(aw.fields.FieldType.FIELD_CITATION, " CITATION Book1 \\l 1033 ", "(Doe, 2018)", field_citation)
        self.assertEqual("Book1", field_citation.source_tag)
        self.assertEqual("1033", field_citation.format_language_id)

        field_bibliography = doc.range.fields[4].as_field_bibliography()

        self.verify_field(aw.fields.FieldType.FIELD_BIBLIOGRAPHY, " BIBLIOGRAPHY ",
            "Cardholder, A. (2018). My Book, Vol. II. New York: Doe Co. Ltd.\rDoe, J. (2018). My Book, Vol I. London: Doe Co. Ltd.\r", field_bibliography)

    def test_field_data(self):

        #ExStart
        #ExFor:FieldData
        #ExSummary:Shows how to insert a DATA field into a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        field = builder.insert_field(aw.fields.FieldType.FIELD_DATA, True).as_field_data()
        self.assertEqual(" DATA ", field.get_field_code())
        #ExEnd

        self.verify_field(aw.fields.FieldType.FIELD_DATA, " DATA ", "", DocumentHelper.save_open(doc).range.fields[0])

    def test_field_include(self):

        #ExStart
        #ExFor:FieldInclude
        #ExFor:FieldInclude.bookmark_name
        #ExFor:FieldInclude.lock_fields
        #ExFor:FieldInclude.source_full_name
        #ExFor:FieldInclude.text_converter
        #ExSummary:Shows how to create an INCLUDE field, and set its properties.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # We can use an INCLUDE field to import a portion of another document in the local file system.
        # The bookmark from the other document that we reference with this field contains this imported portion.
        field = builder.insert_field(aw.fields.FieldType.FIELD_INCLUDE, True).as_field_include()
        field.source_full_name = MY_DIR + "Bookmarks.docx"
        field.bookmark_name = "MyBookmark1"
        field.lock_fields = False
        field.text_converter = "Microsoft Word"

        self.assertRegex(field.get_field_code(), r' INCLUDE .* MyBookmark1 \\c "Microsoft Word"')

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_include.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_include.docx")
        field = doc.range.fields[0].as_field_include()

        self.assertEqual(aw.fields.FieldType.FIELD_INCLUDE, field.type)
        self.assertEqual("First bookmark.", field.result)
        self.assertRegex(field.get_field_code(), r' INCLUDE .* MyBookmark1 \\c "Microsoft Word"')

        self.assertEqual(MY_DIR + "Bookmarks.docx", field.source_full_name)
        self.assertEqual("MyBookmark1", field.bookmark_name)
        self.assertFalse(field.lock_fields)
        self.assertEqual("Microsoft Word", field.text_converter)

    def test_field_include_picture(self):

        #ExStart
        #ExFor:FieldIncludePicture
        #ExFor:FieldIncludePicture.graphic_filter
        #ExFor:FieldIncludePicture.is_linked
        #ExFor:FieldIncludePicture.resize_horizontally
        #ExFor:FieldIncludePicture.resize_vertically
        #ExFor:FieldIncludePicture.source_full_name
        #ExFor:FieldImport
        #ExFor:FieldImport.graphic_filter
        #ExFor:FieldImport.is_linked
        #ExFor:FieldImport.source_full_name
        #ExSummary:Shows how to insert images using IMPORT and INCLUDEPICTURE fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Below are two similar field types that we can use to display images linked from the local file system.
        # 1 -  The INCLUDEPICTURE field:
        field_include_picture = builder.insert_field(aw.fields.FieldType.FIELD_INCLUDE_PICTURE, True).as_field_include_picture()
        field_include_picture.source_full_name = IMAGE_DIR + "Transparent background logo.png"

        self.assertRegex(field_include_picture.get_field_code(), " INCLUDEPICTURE  .*")

        # Apply the PNG32.FLT filter.
        field_include_picture.graphic_filter = "PNG32"
        field_include_picture.is_linked = True
        field_include_picture.resize_horizontally = True
        field_include_picture.resize_vertically = True

        # 2 -  The IMPORT field:
        field_import = builder.insert_field(aw.fields.FieldType.FIELD_IMPORT, True).as_field_import()
        field_import.source_full_name = IMAGE_DIR + "Transparent background logo.png"
        field_import.graphic_filter = "PNG32"
        field_import.is_linked = True

        self.assertRegex(field_import.get_field_code(), r" IMPORT  .* \\c PNG32 \\d")

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_include_picture.docx")
        #ExEnd

        self.assertEqual(IMAGE_DIR + "Transparent background logo.png", field_include_picture.source_full_name)
        self.assertEqual("PNG32", field_include_picture.graphic_filter)
        self.assertTrue(field_include_picture.is_linked)
        self.assertTrue(field_include_picture.resize_horizontally)
        self.assertTrue(field_include_picture.resize_vertically)

        self.assertEqual(IMAGE_DIR + "Transparent background logo.png", field_import.source_full_name)
        self.assertEqual("PNG32", field_import.graphic_filter)
        self.assertTrue(field_import.is_linked)

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_include_picture.docx")

        # The INCLUDEPICTURE fields have been converted into shapes with linked images during loading.
        self.assertEqual(0, doc.range.fields.count)
        self.assertEqual(2, doc.get_child_nodes(aw.NodeType.SHAPE, True).count)

        image = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.assertTrue(image.is_image)
        self.assertIsNone(image.image_data.image_bytes)
        self.assertEqual(IMAGE_DIR + "Transparent background logo.png", image.image_data.source_full_name.replace("%20", " "))

        image = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()

        self.assertTrue(image.is_image)
        self.assertIsNone(image.image_data.image_bytes)
        self.assertEqual(IMAGE_DIR + "Transparent background logo.png", image.image_data.source_full_name.replace("%20", " "))

    #ExStart
    #ExFor:FieldIncludeText
    #ExFor:FieldIncludeText.bookmark_name
    #ExFor:FieldIncludeText.encoding
    #ExFor:FieldIncludeText.lock_fields
    #ExFor:FieldIncludeText.mime_type
    #ExFor:FieldIncludeText.namespace_mappings
    #ExFor:FieldIncludeText.source_full_name
    #ExFor:FieldIncludeText.text_converter
    #ExFor:FieldIncludeText.xpath
    #ExFor:FieldIncludeText.xsl_transformation
    #ExSummary:Shows how to create an INCLUDETEXT field, and set its properties.
    @unittest.skip("WORDSNET-17543") #ExSkip
    def test_field_include_text(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Below are two ways to use INCLUDETEXT fields to display the contents of an XML file in the local file system.
        # 1 -  Perform an XSL transformation on an XML document:
        field_include_text = ExField.create_field_include_text(builder, MY_DIR + "CD collection data.xml", False, "text/xml", "XML", "ISO-8859-1")
        field_include_text.xsl_transformation = MY_DIR + "CD collection XSL transformation.xsl"

        builder.writeln()

        # 2 -  Use an XPath to take specific elements from an XML document:
        field_include_text = ExField.create_field_include_text(builder, MY_DIR + "CD collection data.xml", False, "text/xml", "XML", "ISO-8859-1")
        field_include_text.namespace_mappings = "xmlns:n='myNamespace'"
        field_include_text.xpath = "/catalog/cd/title"

        doc.save(ARTIFACTS_DIR + "Field.field_include_text.docx")
        self._test_field_include_text(aw.Document(ARTIFACTS_DIR + "Field.field_include_text.docx")) #ExSkip

    @staticmethod
    def create_field_include_text(builder: aw.DocumentBuilder, source_full_name: str, lock_fields: bool, mime_type: str, text_converter: str, encoding: str) -> aw.fields.FieldIncludeText:
        """Use a document builder to insert an INCLUDETEXT field with custom properties."""

        field_include_text = builder.insert_field(aw.fields.FieldType.FIELD_INCLUDE_TEXT, True).as_field_include_text()
        field_include_text.source_full_name = source_full_name
        field_include_text.lock_fields = lock_fields
        field_include_text.mime_type = mime_type
        field_include_text.text_converter = text_converter
        field_include_text.encoding = encoding

        return field_include_text

    #ExEnd

    def _test_field_include_text(self, doc: aw.Document):

        doc = DocumentHelper.save_open(doc)

        field_include_text = doc.range.fields[0].as_field_include_text()
        self.assertEqual(MY_DIR + "CD collection data.xml", field_include_text.source_full_name)
        self.assertEqual(MY_DIR + "CD collection XSL transformation.xsl", field_include_text.xsl_transformation)
        self.assertFalse(field_include_text.lock_fields)
        self.assertEqual("text/xml", field_include_text.mime_type)
        self.assertEqual("XML", field_include_text.text_converter)
        self.assertEqual("ISO-8859-1", field_include_text.encoding)
        self.assertEqual(" INCLUDETEXT  \"" + MY_DIR.replace("\\", "\\\\") + "CD collection data.xml\" \\m text/xml \\c XML \\e ISO-8859-1 \\t \"" +
                        MY_DIR.replace("\\", "\\\\") + "CD collection XSL transformation.xsl\"",
            field_include_text.get_field_code())
        self.assertTrue(field_include_text.result.startswith("My CD Collection"))

        cd_collection_data = XmlDocument()
        with open(MY_DIR + "CD collection data.xml", "rt", encoding="utf-8") as file:
            cd_collection_data.load_xml(file.read())

        catalog_data = cd_collection_data.child_nodes[0]

        cd_collection_xsl_transformation = XmlDocument()
        with open(MY_DIR + "CD collection XSL transformation.xsl", "rt", encoding="utf-8") as file:
            cd_collection_xsl_transformation.load_xml(file.read())

        table = doc.first_section.body.tables[0]

        manager = XmlNamespaceManager(cd_collection_xsl_transformation.name_table)
        manager.add_namespace("xsl", "http://www.w3.org/1999/XSL/Transform")

        for i in range(table.rows.count):
            for j in range(table.rows[i].count):
                if i == 0:
                    # When on the first row from the input document's table, ensure that all table's cells match all XML element Names.
                    for k in range(table.rows.count - 1):
                        self.assertEqual(catalog_data.child_nodes[k].child_nodes[j].name,
                            table.rows[i].cells[j].get_text().replace(aw.ControlChar.CELL, "").lower())

                    # Also, make sure that the whole first row has the same color as the XSL transform.
                    self.assertEqual(cd_collection_xsl_transformation.select_nodes("//xsl:stylesheet/xsl:template/html/body/table/tr", manager)[0].attributes.get_named_item("bgcolor").value,
                        drawing.ColorTranslator.to_html(table.rows[i].cells[j].cell_format.shading.background_pattern_color).lower())
                else:
                    # When on all other rows of the input document's table, ensure that cell contents match XML element Values.
                    self.assertEqual(catalog_data.child_nodes[i - 1].child_nodes[j].first_child.value,
                        table.rows[i].cells[j].get_text().replace(aw.ControlChar.CELL, ""))
                    self.assertEqual(drawing.Color.empty(), table.rows[i].cells[j].cell_format.shading.background_pattern_color)

                self.assertEqual(
                    float(cd_collection_xsl_transformation.select_nodes("//xsl:stylesheet/xsl:template/html/body/table", manager)[0].attributes.get_named_item("border").value) * 0.75,
                    table.first_row.row_format.borders.bottom.line_width)

        field_include_text = doc.range.fields[1].as_field_include_text()
        self.assertEqual(MY_DIR + "CD collection data.xml", field_include_text.source_full_name)
        self.assertIsNone(field_include_text.xsl_transformation)
        self.assertFalse(field_include_text.lock_fields)
        self.assertEqual("text/xml", field_include_text.mime_type)
        self.assertEqual("XML", field_include_text.text_converter)
        self.assertEqual("ISO-8859-1", field_include_text.encoding)
        self.assertEqual(" INCLUDETEXT  \"" + MY_DIR.replace("\\", "\\\\") + "CD collection data.xml\" \\m text/xml \\c XML \\e ISO-8859-1 \\n xmlns:n='myNamespace' \\x /catalog/cd/title",
            field_include_text.get_field_code())

        expected_field_result = ""
        for i in range(catalog_data.child_nodes.count):
            expected_field_result += catalog_data.child_nodes[i].child_nodes[0].child_nodes[0].value

        self.assertEqual(expected_field_result, field_include_text.result)

    @unittest.skip("WORDSNET-17545")
    def test_field_hyperlink(self):

        #ExStart
        #ExFor:FieldHyperlink
        #ExFor:FieldHyperlink.address
        #ExFor:FieldHyperlink.is_image_map
        #ExFor:FieldHyperlink.open_in_new_window
        #ExFor:FieldHyperlink.screen_tip
        #ExFor:FieldHyperlink.sub_address
        #ExFor:FieldHyperlink.target
        #ExSummary:Shows how to use HYPERLINK fields to link to documents in the local file system.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        field = builder.insert_field(aw.fields.FieldType.FIELD_HYPERLINK, True).as_field_hyperlink()

        # When we click this HYPERLINK field in Microsoft Word,
        # it will open the linked document and then place the cursor at the specified bookmark.
        field.address = MY_DIR + "Bookmarks.docx"
        field.sub_address = "MyBookmark3"
        field.screen_tip = "Open " + field.address + " on bookmark " + field.sub_address + " in a new window"

        builder.writeln()

        # When we click this HYPERLINK field in Microsoft Word,
        # it will open the linked document, and automatically scroll down to the specified iframe.
        field = builder.insert_field(aw.fields.FieldType.FIELD_HYPERLINK, True).as_field_hyperlink()
        field.address = MY_DIR + "Iframes.html"
        field.screen_tip = "Open " + field.address
        field.target = "iframe_3"
        field.open_in_new_window = True
        field.is_image_map = False

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_hyperlink.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_hyperlink.docx")
        field = doc.range.fields[0].as_field_hyperlink()

        self.verify_field(aw.fields.FieldType.FIELD_HYPERLINK,
            " HYPERLINK \"" + MY_DIR.replace("\\", "\\\\") + "Bookmarks.docx\" \\l \"MyBookmark3\" \\o \"Open " + MY_DIR + "Bookmarks.docx on bookmark MyBookmark3 in a new window\" ",
            MY_DIR + "Bookmarks.docx - MyBookmark3", field)
        self.assertEqual(MY_DIR + "Bookmarks.docx", field.address)
        self.assertEqual("MyBookmark3", field.sub_address)
        self.assertEqual("Open " + field.address.replace("\\", "") + " on bookmark " + field.sub_address + " in a new window", field.screen_tip)

        field = doc.range.fields[1].as_field_hyperlink()

        self.verify_field(aw.fields.FieldType.FIELD_HYPERLINK, " HYPERLINK \"file:///" + MY_DIR.replace("\\", "\\\\").replace(" ", "%20") + "Iframes.html\" \\t \"iframe_3\" \\o \"Open " + MY_DIR.replace("\\", "\\\\") + "Iframes.html\" ",
            MY_DIR + "Iframes.html", field)
        self.assertEqual("file:///" + MY_DIR.replace(" ", "%20") + "Iframes.html", field.address)
        self.assertEqual("Open " + MY_DIR + "Iframes.html", field.screen_tip)
        self.assertEqual("iframe_3", field.target)
        self.assertFalse(field.open_in_new_window)
        self.assertFalse(field.is_image_map)

    ##ExStart
    ##ExFor:MergeFieldImageDimension
    ##ExFor:MergeFieldImageDimension.__init__
    ##ExFor:MergeFieldImageDimension.__init__(float)
    ##ExFor:MergeFieldImageDimension.__init__(float,MergeFieldImageDimensionUnit)
    ##ExFor:MergeFieldImageDimension.unit
    ##ExFor:MergeFieldImageDimension.value
    ##ExFor:MergeFieldImageDimensionUnit
    ##ExFor:ImageFieldMergingArgs
    ##ExFor:ImageFieldMergingArgs.image_file_name
    ##ExFor:ImageFieldMergingArgs.image_width
    ##ExFor:ImageFieldMergingArgs.image_height
    ##ExSummary:Shows how to set the dimensions of images as MERGEFIELDS accepts them during a mail merge.
    #def test_merge_field_image_dimension(self):

    #    doc = aw.Document()

    #    # Insert a MERGEFIELD that will accept images from a source during a mail merge. Use the field code to reference
    #    # a column in the data source containing local system filenames of images we wish to use in the mail merge.
    #    builder = aw.DocumentBuilder(doc)
    #    field = builder.insert_field("MERGEFIELD Image:ImageColumn").as_field_merge_field()

    #    # The data source should have such a column named "ImageColumn".
    #    self.assertEqual("Image:ImageColumn", field.field_name)

    #    # Create a suitable data source.
    #    data_table = DataTable("Images")
    #    data_table.columns.add(DataColumn("ImageColumn"))
    #    data_table.rows.add(IMAGE_DIR + "Logo.jpg")
    #    data_table.rows.add(IMAGE_DIR + "Transparent background logo.png")
    #    data_table.rows.add(IMAGE_DIR + "Enhanced Windows MetaFile.emf")

    #    # Configure a callback to modify the sizes of images at merge time, then execute the mail merge.
    #    doc.mail_merge.field_merging_callback = ExField.MergedImageResizer(200, 200, aw.fields.MergeFieldImageDimensionUnit.POINT)
    #    doc.mail_merge.execute(data_table)

    #    doc.update_fields()
    #    doc.save(ARTIFACTS_DIR + "Field.mergefield.image_dimension.docx")
    #    self._test_merge_field_image_dimension(doc) #ExSkip

    #class MergedImageResizer(aw.mailmerging.IFieldMergingCallback):
    #    """Sets the size of all mail merged images to one defined width and height."""

    #    def __init__(self, image_width: float, image_height: float, unit: aw.fields.MergeFieldImageDimensionUnit):

    #        self.image_width = image_width
    #        self.image_height = image_height
    #        self.unit = unit

    #    def field_merging(self, e: aw.mailmerging.FieldMergingArgs):

    #        raise NotImplementedError()

    #    def image_field_merging(self, args: aw.mailmerging.ImageFieldMergingArgs):

    #        args.image_file_name = args.field_value.to_string()
    #        args.image_width = aw.fields.MergeFieldImageDimension(self.image_width, self.unit)
    #        args.image_height = aw.fields.MergeFieldImageDimension(self.image_height, self.unit)

    #        self.assertEqual(self.image_width, args.image_width.value)
    #        self.assertEqual(mUnit, args.image_width.unit)
    #        self.assertEqual(self.image_height, args.image_height.value)
    #        self.assertEqual(mUnit, args.image_height.unit)

    ##ExEnd

    def _test_merge_field_image_dimension(self, doc: aw.Document):

        doc = DocumentHelper.save_open(doc)

        self.assertEqual(0, doc.range.fields.count)
        self.assertEqual(3, doc.get_child_nodes(aw.NodeType.SHAPE, True).count)

        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, shape)
        self.assertEqual(200.0, shape.width)
        self.assertEqual(200.0, shape.height)

        shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()

        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.PNG, shape)
        self.assertEqual(200.0, shape.width)
        self.assertEqual(200.0, shape.height)

        shape = doc.get_child(aw.NodeType.SHAPE, 2, True).as_shape()

        self.verify_image_in_shape(534, 534, aw.drawing.ImageType.EMF, shape)
        self.assertEqual(200.0, shape.width)
        self.assertEqual(200.0, shape.height)

    ##ExStart
    ##ExFor:ImageFieldMergingArgs.image
    ##ExSummary:Shows how to use a callback to customize image merging logic.
    #def test_merge_field_images(self):

    #    doc = aw.Document()

    #    # Insert a MERGEFIELD that will accept images from a source during a mail merge. Use the field code to reference
    #    # a column in the data source which contains local system filenames of images we wish to use in the mail merge.
    #    builder = aw.DocumentBuilder(doc)
    #    field = builder.insert_field("MERGEFIELD Image:ImageColumn").as_field_merge_field()

    #    # In this case, the field expects the data source to have such a column named "ImageColumn".
    #    self.assertEqual("Image:ImageColumn", field.field_name)

    #    # Filenames can be lengthy, and if we can find a way to avoid storing them in the data source,
    #    # we may considerably reduce its size.
    #    # Create a data source that refers to images using short names.
    #    data_table = DataTable("Images")
    #    data_table.columns.add(DataColumn("ImageColumn"))
    #    data_table.rows.add("Dark logo")
    #    data_table.rows.add("Transparent logo")

    #    # Assign a merging callback that contains all logic that processes those names,
    #    # and then execute the mail merge.
    #    doc.mail_merge.field_merging_callback = ExField.ImageFilenameCallback()
    #    doc.mail_merge.execute(data_table)

    #    doc.save(ARTIFACTS_DIR + "Field.mergefield.images.docx")
    #    self._test_merge_field_images(aw.Document(ARTIFACTS_DIR + "Field.mergefield.images.docx")) #ExSkip

    #class ImageFilenameCallback(ex.fields.IFieldMergingCallback):
    #    """Contains a dictionary that maps names of images to local system filenames that contain these images.
    #    If a mail merge data source uses one of the dictionary's names to refer to an image,
    #    this callback will pass the respective filename to the merge destination."""

    #    def __init__(self):

    #        self.image_filenames: Dict[str, str] = {}
    #        self.image_filenames["Dark logo"] = IMAGE_DIR + "Logo.jpg"
    #        self.image_filenames["Transparent logo"] = IMAGE_DIR + "Transparent background logo.png"

    #    def ield_merging(self, args: aw.fields.FieldMergingArgs):

    #        raise NotImplementedError()

    #    def image_field_merging(self, args: aw.fields.ImageFieldMergingArgs):

    #        if self.image_filenames.contains_key(args.field_value.to_string()):
    #            args.image = drawing.Image.from_file(self.image_filenames[args.field_value.to_string()])

    #        self.assertIsNotNone(args.image)

    ##ExEnd

    def _test_merge_field_images(self, doc: aw.Document):

        doc = DocumentHelper.save_open(doc)

        self.assertEqual(0, doc.range.fields.count)
        self.assertEqual(2, doc.get_child_nodes(aw.NodeType.SHAPE, True).count)

        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.JPEG, shape)
        self.assertEqual(300.0, shape.width)
        self.assertEqual(300.0, shape.height)

        shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()

        self.verify_image_in_shape(400, 400, aw.drawing.ImageType.PNG, shape)
        self.assertEqual(300.0, shape.width, 1)
        self.assertEqual(300.0, shape.height, 1)

    @unittest.skip("WORDSNET-17524")
    def test_field_index_filter(self):

        #ExStart
        #ExFor:FieldIndex
        #ExFor:FieldIndex.bookmark_name
        #ExFor:FieldIndex.entry_type
        #ExFor:FieldXE
        #ExFor:FieldXE.entry_type
        #ExFor:FieldXE.text
        #ExSummary:Shows how to create an INDEX field, and then use XE fields to populate it with entries.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create an INDEX field which will display an entry for each XE field found in the document.
        # Each entry will display the XE field's Text property value on the left side
        # and the page containing the XE field on the right.
        # If the XE fields have the same value in their "text" property,
        # the INDEX field will group them into one entry.
        index = builder.insert_field(aw.fields.FieldType.FIELD_INDEX, True).as_field_index()

        # Configure the INDEX field only to display XE fields that are within the bounds
        # of a bookmark named "MainBookmark", and whose "entry_type" properties have a value of "A".
        # For both INDEX and XE fields, the "entry_type" property only uses the first character of its string value.
        index.bookmark_name = "MainBookmark"
        index.entry_type = "A"

        self.assertEqual(" INDEX  \\b MainBookmark \\f A", index.get_field_code())

        # On a new page, start the bookmark with a name that matches the value
        # of the INDEX field's "bookmark_name" property.
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.start_bookmark("MainBookmark")

        # The INDEX field will pick up this entry because it is inside the bookmark,
        # and its entry type also matches the INDEX field's entry type.
        index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
        index_entry.text = "Index entry 1"
        index_entry.entry_type = "A"

        self.assertEqual(" XE  \"Index entry 1\" \\f A", index_entry.get_field_code())

        # Insert an XE field that will not appear in the INDEX because the entry types do not match.
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
        index_entry.text = "Index entry 2"
        index_entry.entry_type = "B"

        # End the bookmark and insert an XE field afterwards.
        # It is of the same type as the INDEX field, but will not appear
        # since it is outside the bookmark's boundaries.
        builder.end_bookmark("MainBookmark")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
        index_entry.text = "Index entry 3"
        index_entry.entry_type = "A"

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_index_filter.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_index_filter.docx")
        index = doc.range.fields[0].as_field_index()

        self.verify_field(aw.fields.FieldType.FIELD_INDEX, " INDEX  \\b MainBookmark \\f A", "Index entry 1, 2\r", index)
        self.assertEqual("MainBookmark", index.bookmark_name)
        self.assertEqual("A", index.entry_type)

        index_entry = doc.range.fields[1].as_field_xe()

        self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " XE  \"Index entry 1\" \\f A", "", index_entry)
        self.assertEqual("Index entry 1", index_entry.text)
        self.assertEqual("A", index_entry.entry_type)

        index_entry = doc.range.fields[2].as_field_xe()

        self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " XE  \"Index entry 2\" \\f B", "", index_entry)
        self.assertEqual("Index entry 2", index_entry.text)
        self.assertEqual("B", index_entry.entry_type)

        index_entry = doc.range.fields[3].as_field_xe()

        self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " XE  \"Index entry 3\" \\f A", "", index_entry)
        self.assertEqual("Index entry 3", index_entry.text)
        self.assertEqual("A", index_entry.entry_type)

    @unittest.skip("WORDSNET-17524")
    def test_field_index_formatting(self):

        #ExStart
        #ExFor:FieldIndex
        #ExFor:FieldIndex.heading
        #ExFor:FieldIndex.number_of_columns
        #ExFor:FieldIndex.language_id
        #ExFor:FieldIndex.letter_range
        #ExFor:FieldXE
        #ExFor:FieldXE.is_bold
        #ExFor:FieldXE.is_italic
        #ExFor:FieldXE.text
        #ExSummary:Shows how to populate an INDEX field with entries using XE fields, and also modify its appearance.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create an INDEX field which will display an entry for each XE field found in the document.
        # Each entry will display the XE field's Text property value on the left side,
        # and the number of the page that contains the XE field on the right.
        # If the XE fields have the same value in their "text" property,
        # the INDEX field will group them into one entry.
        index = builder.insert_field(aw.fields.FieldType.FIELD_INDEX, True).as_field_index()
        index.language_id = "1033" # en-US

        # Setting this property's value to "A" will group all the entries by their first letter,
        # and place that letter in uppercase above each group.
        index.heading = "A"

        # Set the table created by the INDEX field to span over 2 columns.
        index.number_of_columns = "2"

        # Set any entries with starting letters outside the "a-c" character range to be omitted.
        index.letter_range = "a-c"

        self.assertEqual(" INDEX  \\z 1033 \\h A \\c 2 \\p a-c", index.get_field_code())

        # These next two XE fields will show up under the "A" heading,
        # with their respective text stylings also applied to their page numbers.
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
        index_entry.text = "Apple"
        index_entry.is_italic = True

        self.assertEqual(" XE  Apple \\i", index_entry.get_field_code())

        builder.insert_break(aw.BreakType.PAGE_BREAK)
        index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
        index_entry.text = "Apricot"
        index_entry.is_bold = True

        self.assertEqual(" XE  Apricot \\b", index_entry.get_field_code())

        # Both the next two XE fields will be under a "B" and "C" heading in the INDEX fields table of contents.
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
        index_entry.text = "Banana"

        builder.insert_break(aw.BreakType.PAGE_BREAK)
        index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
        index_entry.text = "Cherry"

        # INDEX fields sort all entries alphabetically, so this entry will show up under "A" with the other two.
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
        index_entry.text = "Avocado"

        # This entry will not appear because it starts with the letter "D",
        # which is outside the "a-c" character range that the INDEX field's LetterRange property defines.
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
        index_entry.text = "Durian"

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_index_formatting.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_index_formatting.docx")
        index = doc.range.fields[0].as_field_index()

        self.assertEqual("1033", index.language_id)
        self.assertEqual("A", index.heading)
        self.assertEqual("2", index.number_of_columns)
        self.assertEqual("a-c", index.letter_range)
        self.assertEqual(" INDEX  \\z 1033 \\h A \\c 2 \\p a-c", index.get_field_code())
        self.assertEqual("\fA\r" +
                        "Apple, 2\r" +
                        "Apricot, 3\r" +
                        "Avocado, 6\r" +
                        "B\r" +
                        "Banana, 4\r" +
                        "C\r" +
                        "Cherry, 5\r\f", index.result)

        index_entry = doc.range.fields[1].as_field_xe()

        self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " XE  Apple \\i", "", index_entry)
        self.assertEqual("Apple", index_entry.text)
        self.assertFalse(index_entry.is_bold)
        self.assertTrue(index_entry.is_italic)

        index_entry = doc.range.fields[2].as_field_xe()

        self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " XE  Apricot \\b", "", index_entry)
        self.assertEqual("Apricot", index_entry.text)
        self.assertTrue(index_entry.is_bold)
        self.assertFalse(index_entry.is_italic)

        index_entry = doc.range.fields[3].as_field_xe()

        self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " XE  Banana", "", index_entry)
        self.assertEqual("Banana", index_entry.text)
        self.assertFalse(index_entry.is_bold)
        self.assertFalse(index_entry.is_italic)

        index_entry = doc.range.fields[4].as_field_xe()

        self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " XE  Cherry", "", index_entry)
        self.assertEqual("Cherry", index_entry.text)
        self.assertFalse(index_entry.is_bold)
        self.assertFalse(index_entry.is_italic)

        index_entry = doc.range.fields[5].as_field_xe()

        self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " XE  Avocado", "", index_entry)
        self.assertEqual("Avocado", index_entry.text)
        self.assertFalse(index_entry.is_bold)
        self.assertFalse(index_entry.is_italic)

        index_entry = doc.range.fields[6].as_field_xe()

        self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " XE  Durian", "", index_entry)
        self.assertEqual("Durian", index_entry.text)
        self.assertFalse(index_entry.is_bold)
        self.assertFalse(index_entry.is_italic)

    @unittest.skip("WORDSNET-17524")
    def test_field_index_sequence(self):

        #ExStart
        #ExFor:FieldIndex.has_sequence_name
        #ExFor:FieldIndex.sequence_name
        #ExFor:FieldIndex.sequence_separator
        #ExSummary:Shows how to split a document into portions by combining INDEX and SEQ fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create an INDEX field which will display an entry for each XE field found in the document.
        # Each entry will display the XE field's Text property value on the left side,
        # and the number of the page that contains the XE field on the right.
        # If the XE fields have the same value in their "text" property,
        # the INDEX field will group them into one entry.
        index = builder.insert_field(aw.fields.FieldType.FIELD_INDEX, True).as_field_index()

        # In the sequence_name property, name a SEQ field sequence. Each entry of this INDEX field will now also display
        # the number that the sequence count is on at the XE field location that created this entry.
        index.sequence_name = "MySequence"

        # Set text that will around the sequence and page numbers to explain their meaning to the user.
        # An entry created with this configuration will display something like "MySequence at 1 on page 1" at its page number.
        # "page_number_separator" and "sequence_separator" cannot be longer than 15 characters.
        index.page_number_separator = "\tMySequence at "
        index.sequence_separator = " on page "
        self.assertTrue(index.has_sequence_name)

        self.assertEqual(" INDEX  \\s MySequence \\e \"\tMySequence at \" \\d \" on page \"", index.get_field_code())

        # SEQ fields display a count that increments at each SEQ field.
        # These fields also maintain separate counts for each unique named sequence
        # identified by the SEQ field's "sequence_identifier" property.
        # Insert a SEQ field which moves the "MySequence" sequence to 1.
        # This field no different from normal document text. It will not appear on an INDEX field's table of contents.
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        sequence_field = builder.insert_field(aw.fields.FieldType.FIELD_SEQUENCE, True).as_field_seq()
        sequence_field.sequence_identifier = "MySequence"

        self.assertEqual(" SEQ  MySequence", sequence_field.get_field_code())

        # Insert an XE field which will create an entry in the INDEX field.
        # Since "MySequence" is at 1 and this XE field is on page 2, along with the custom separators we defined above,
        # this field's INDEX entry will display "Cat" on the left side, and "MySequence at 1 on page 2" on the right.
        index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
        index_entry.text = "Cat"

        self.assertEqual(" XE  Cat", index_entry.get_field_code())

        # Insert a page break and use SEQ fields to advance "MySequence" to 3.
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        sequence_field = builder.insert_field(aw.fields.FieldType.FIELD_SEQUENCE, True).as_field_seq()
        sequence_field.sequence_identifier = "MySequence"
        sequence_field = builder.insert_field(aw.fields.FieldType.FIELD_SEQUENCE, True).as_field_seq()
        sequence_field.sequence_identifier = "MySequence"

        # Insert an XE field with the same "text" property as the one above.
        # The INDEX entry will group XE fields with matching values in the "text" property
        # into one entry as opposed to making an entry for each XE field.
        # Since we are on page 2 with "MySequence" at 3, ", 3 on page 3" will be appended to the same INDEX entry as above.
        # The page number portion of that INDEX entry will now display "MySequence at 1 on page 2, 3 on page 3".
        index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
        index_entry.text = "Cat"

        # Insert an XE field with a new and unique "text" property value.
        # This will add a new entry, with MySequence at 3 on page 4.
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
        index_entry.text = "Dog"

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_index_sequence.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_index_sequence.docx")
        index = doc.range.fields[0].as_field_index()

        self.assertEqual("MySequence", index.sequence_name)
        self.assertEqual("\tMySequence at ", index.page_number_separator)
        self.assertEqual(" on page ", index.sequence_separator)
        self.assertTrue(index.has_sequence_name)
        self.assertEqual(" INDEX  \\s MySequence \\e \"\tMySequence at \" \\d \" on page \"", index.get_field_code())
        self.assertEqual("Cat\tMySequence at 1 on page 2, 3 on page 3\r" +
                        "Dog\tMySequence at 3 on page 4\r", index.result)

        self.assertEqual(3, len([f for f in doc.range.fields if f.type == aw.fields.FieldType.FIELD_SEQUENCE]))

    @unittest.skip("WORDSNET-17524")
    def test_field_index_page_number_separator(self):

        #ExStart
        #ExFor:FieldIndex.has_page_number_separator
        #ExFor:FieldIndex.page_number_separator
        #ExFor:FieldIndex.page_number_list_separator
        #ExSummary:Shows how to edit the page number separator in an INDEX field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create an INDEX field which will display an entry for each XE field found in the document.
        # Each entry will display the XE field's "text" property value on the left side,
        # and the number of the page that contains the XE field on the right.
        # The INDEX entry will group XE fields with matching values in the "text" property
        # into one entry as opposed to making an entry for each XE field.
        index = builder.insert_field(aw.fields.FieldType.FIELD_INDEX, True).as_field_index()

        # If our INDEX field has an entry for a group of XE fields,
        # this entry will display the number of each page that contains an XE field that belongs to this group.
        # We can set custom separators to customize the appearance of these page numbers.
        index.page_number_separator = ", on page(s) "
        index.page_number_list_separator = " & "

        self.assertEqual(" INDEX  \\e \", on page(s) \" \\l \" & \"", index.get_field_code())
        self.assertTrue(index.has_page_number_separator)

        # After we insert these XE fields, the INDEX field will display "First entry, on page(s) 2 & 3 & 4".
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
        index_entry.text = "First entry"

        self.assertEqual(" XE  \"First entry\"", index_entry.get_field_code())

        builder.insert_break(aw.BreakType.PAGE_BREAK)
        index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
        index_entry.text = "First entry"

        builder.insert_break(aw.BreakType.PAGE_BREAK)
        index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
        index_entry.text = "First entry"

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_index_page_number_separator.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_index_page_number_separator.docx")
        index = doc.range.fields[0].as_field_index()

        self.verify_field(aw.fields.FieldType.FIELD_INDEX, " INDEX  \\e \", on page(s) \" \\l \" & \"", "First entry, on page(s) 2 & 3 & 4\r", index)
        self.assertEqual(", on page(s) ", index.page_number_separator)
        self.assertEqual(" & ", index.page_number_list_separator)
        self.assertTrue(index.has_page_number_separator)

    @unittest.skip("WORDSNET-17524")
    def test_field_index_page_range_bookmark(self):

        #ExStart
        #ExFor:FieldIndex.page_range_separator
        #ExFor:FieldXE.has_page_range_bookmark_name
        #ExFor:FieldXE.page_range_bookmark_name
        #ExSummary:Shows how to specify a bookmark's spanned pages as a page range for an INDEX field entry.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create an INDEX field which will display an entry for each XE field found in the document.
        # Each entry will display the XE field's "text" property value on the left side,
        # and the number of the page that contains the XE field on the right.
        # The INDEX entry will collect all XE fields with matching values in the "text" property
        # into one entry as opposed to making an entry for each XE field.
        index = builder.insert_field(aw.fields.FieldType.FIELD_INDEX, True).as_field_index()

        # For INDEX entries that display page ranges, we can specify a separator string
        # which will appear between the number of the first page, and the number of the last.
        index.page_number_separator = ", on page(s) "
        index.page_range_separator = " to "

        self.assertEqual(" INDEX  \\e \", on page(s) \" \\g \" to \"", index.get_field_code())

        builder.insert_break(aw.BreakType.PAGE_BREAK)
        index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
        index_entry.text = "My entry"

        # If an XE field names a bookmark using the page_range_bookmark_name property,
        # its INDEX entry will show the range of pages that the bookmark spans
        # instead of the number of the page that contains the XE field.
        index_entry.page_range_bookmark_name = "MyBookmark"

        self.assertEqual(" XE  \"My entry\" \\r MyBookmark", index_entry.get_field_code())
        self.assertTrue(index_entry.has_page_range_bookmark_name)

        # Insert a bookmark that starts on page 3 and ends on page 5.
        # The INDEX entry for the XE field that references this bookmark will display this page range.
        # In our table, the INDEX entry will display "My entry, on page(s) 3 to 5".
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.start_bookmark("MyBookmark")
        builder.write("Start of MyBookmark")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.write("End of MyBookmark")
        builder.end_bookmark("MyBookmark")

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_index_page_range_bookmark.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_index_page_range_bookmark.docx")
        index = doc.range.fields[0].as_field_index()

        self.verify_field(aw.fields.FieldType.FIELD_INDEX, " INDEX  \\e \", on page(s) \" \\g \" to \"", "My entry, on page(s) 3 to 5\r", index)
        self.assertEqual(", on page(s) ", index.page_number_separator)
        self.assertEqual(" to ", index.page_range_separator)

        index_entry = doc.range.fields[1].as_field_xe()

        self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " XE  \"My entry\" \\r MyBookmark", "", index_entry)
        self.assertEqual("My entry", index_entry.text)
        self.assertEqual("MyBookmark", index_entry.page_range_bookmark_name)
        self.assertTrue(index_entry.has_page_range_bookmark_name)

    @unittest.skip("WORDSNET-17524")
    def test_field_index_cross_reference_separator(self):

        #ExStart
        #ExFor:FieldIndex.cross_reference_separator
        #ExFor:FieldXE.page_number_replacement
        #ExSummary:Shows how to define cross references in an INDEX field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create an INDEX field which will display an entry for each XE field found in the document.
        # Each entry will display the XE field's "text" property value on the left side,
        # and the number of the page that contains the XE field on the right.
        # The INDEX entry will collect all XE fields with matching values in the "text" property
        # into one entry as opposed to making an entry for each XE field.
        index = builder.insert_field(aw.fields.FieldType.FIELD_INDEX, True).as_field_index()

        # We can configure an XE field to get its INDEX entry to display a string instead of a page number.
        # First, for entries that substitute a page number with a string,
        # specify a custom separator between the XE field's "text" property value and the string.
        index.cross_reference_separator = ", see: "

        self.assertEqual(" INDEX  \\k \", see: \"", index.get_field_code())

        # Insert an XE field, which creates a regular INDEX entry which displays this field's page number,
        # and does not invoke the CrossReferenceSeparator value.
        # The entry for this XE field will display "Apple, 2".
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
        index_entry.text = "Apple"

        self.assertEqual(" XE  Apple", index_entry.get_field_code())

        # Insert another XE field on page 3 and set a value for the "page_number_replacement" property.
        # This value will show up instead of the number of the page that this field is on,
        # and the INDEX field's CrossReferenceSeparator value will appear in front of it.
        # The entry for this XE field will display "Banana, see: Tropical fruit".
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
        index_entry.text = "Banana"
        index_entry.page_number_replacement = "Tropical fruit"

        self.assertEqual(" XE  Banana \\t \"Tropical fruit\"", index_entry.get_field_code())

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_index_cross_reference_separator.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_index_cross_reference_separator.docx")
        index = doc.range.fields[0].as_field_index()

        self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " INDEX  \\k \", see: \"",
            "Apple, 2\r" +
            "Banana, see: Tropical fruit\r", index)
        self.assertEqual(", see: ", index.cross_reference_separator)

        index_entry = doc.range.fields[1].as_field_xe()

        self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " XE  Apple", "", index_entry)
        self.assertEqual("Apple", index_entry.text)
        self.assertIsNone(index_entry.page_number_replacement)

        index_entry = doc.range.fields[2].as_field_xe()

        self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " XE  Banana \\t \"Tropical fruit\"", "", index_entry)
        self.assertEqual("Banana", index_entry.text)
        self.assertEqual("Tropical fruit", index_entry.page_number_replacement)

    @unittest.skip("WORDSNET-17524")
    def test_field_index_subheading(self):

        for run_subentries_on_the_same_line in (True, False):
            with self.subTest(run_subentries_on_the_same_line=run_subentries_on_the_same_line):
                #ExStart
                #ExFor:FieldIndex.run_subentries_on_same_line
                #ExSummary:Shows how to work with subentries in an INDEX field.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Create an INDEX field which will display an entry for each XE field found in the document.
                # Each entry will display the XE field's "text" property value on the left side,
                # and the number of the page that contains the XE field on the right.
                # The INDEX entry will collect all XE fields with matching values in the "text" property
                # into one entry as opposed to making an entry for each XE field.
                index = builder.insert_field(aw.fields.FieldType.FIELD_INDEX, True).as_field_index()
                index.page_number_separator = ", see page "
                index.heading = "A"

                # XE fields that have a "text" property whose value becomes the heading of the INDEX entry.
                # If this value contains two string segments split by a colon (the INDEX entry will treat :) delimiter,
                # the first segment is heading, and the second segment will become the subheading.
                # The INDEX field first groups entries alphabetically, then, if there are multiple XE fields with the same
                # headings, the INDEX field will further subgroup them by the values of these headings.
                # There can be multiple subgrouping layers, depending on how many times
                # the "text" properties of XE fields get segmented like this.
                # By default, an INDEX field entry group will create a new line for every subheading within this group.
                # We can set the "run_subentries_on_same_line" flag to "True" to keep the heading,
                # and every subheading for the group on one line instead, which will make the INDEX field more compact.
                index.run_subentries_on_same_line = run_subentries_on_the_same_line

                if run_subentries_on_the_same_line:
                    self.assertEqual(" INDEX  \\e \", see page \" \\h A \\r", index.get_field_code())
                else:
                    self.assertEqual(" INDEX  \\e \", see page \" \\h A", index.get_field_code())

                # Insert two XE fields, each on a new page, and with the same heading named "Heading 1",
                # which the INDEX field will use to group them.
                # If "run_subentries_on_same_line" is "False", then the INDEX table will create three lines:
                # one line for the grouping heading "Heading 1", and one more line for each subheading.
                # If "run_subentries_on_same_line" is "True", then the INDEX table will create a one-line
                # entry that encompasses the heading and every subheading.
                builder.insert_break(aw.BreakType.PAGE_BREAK)
                index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
                index_entry.text = "Heading 1:Subheading 1"

                self.assertEqual(" XE  \"Heading 1:Subheading 1\"", index_entry.get_field_code())

                builder.insert_break(aw.BreakType.PAGE_BREAK)
                index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
                index_entry.text = "Heading 1:Subheading 2"

                doc.update_fields()
                doc.save(ARTIFACTS_DIR + "Field.field_index_subheading.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "Field.field_index_subheading.docx")
                index = doc.range.fields[0].as_field_index()

                if run_subentries_on_the_same_line:
                    self.verify_field(aw.fields.FieldType.FIELD_INDEX, " INDEX  \\r \\e \", see page \" \\h A",
                        "H\r" +
                        "Heading 1: Subheading 1, see page 2; Subheading 2, see page 3\r", index)
                    self.assertTrue(index.run_subentries_on_same_line)
                else:
                    self.verify_field(aw.fields.FieldType.FIELD_INDEX, " INDEX  \\e \", see page \" \\h A",
                        "H\r" +
                        "Heading 1\r" +
                        "Subheading 1, see page 2\r" +
                        "Subheading 2, see page 3\r", index)
                    self.assertFalse(index.run_subentries_on_same_line)

                index_entry = doc.range.fields[1].as_field_xe()

                self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " XE  \"Heading 1:Subheading 1\"", "", index_entry)
                self.assertEqual("Heading 1:Subheading 1", index_entry.text)

                index_entry = doc.range.fields[2].as_field_xe()

                self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " XE  \"Heading 1:Subheading 2\"", "", index_entry)
                self.assertEqual("Heading 1:Subheading 2", index_entry.text)

    @unittest.skip("WORDSNET-17524")
    def test_field_index_yomi(self):

        for sort_entries_using_yomi in (True, False):
            with self.subTest(sort_entries_using_yomi=sort_entries_using_yomi):
                #ExStart
                #ExFor:FieldIndex.use_yomi
                #ExFor:FieldXE.yomi
                #ExSummary:Shows how to sort INDEX field entries phonetically.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Create an INDEX field which will display an entry for each XE field found in the document.
                # Each entry will display the XE field's "text" property value on the left side,
                # and the number of the page that contains the XE field on the right.
                # The INDEX entry will collect all XE fields with matching values in the "text" property
                # into one entry as opposed to making an entry for each XE field.
                index = builder.insert_field(aw.fields.FieldType.FIELD_INDEX, True).as_field_index()

                # The INDEX table automatically sorts its entries by the values of their "text" properties in alphabetic order.
                # Set the INDEX table to sort entries phonetically using Hiragana instead.
                index.use_yomi = sort_entries_using_yomi

                if sort_entries_using_yomi:
                    self.assertEqual(" INDEX  \\y", index.get_field_code())
                else:
                    self.assertEqual(" INDEX ", index.get_field_code())

                # Insert 4 XE fields, which would show up as entries in the INDEX field's table of contents.
                # The "text" property may contain a word's spelling in Kanji, whose pronunciation may be ambiguous,
                # while the "Yomi" version of the word will spell exactly how it is pronounced using Hiragana.
                # If we set our INDEX field to use Yomi, it will sort these entries
                # by the value of their "yomi" properties, instead of their "text" values.
                builder.insert_break(aw.BreakType.PAGE_BREAK)
                index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
                index_entry.text = "愛子"
                index_entry.yomi = "あ"

                self.assertEqual(" XE  愛子 \\y あ", index_entry.get_field_code())

                builder.insert_break(aw.BreakType.PAGE_BREAK)
                index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
                index_entry.text = "明美"
                index_entry.yomi = "あ"

                builder.insert_break(aw.BreakType.PAGE_BREAK)
                index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
                index_entry.text = "恵美"
                index_entry.yomi = "え"

                builder.insert_break(aw.BreakType.PAGE_BREAK)
                index_entry = builder.insert_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, True).as_field_xe()
                index_entry.text = "愛美"
                index_entry.yomi = "え"

                doc.update_fields()
                doc.save(ARTIFACTS_DIR + "Field.field_index_yomi.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "Field.field_index_yomi.docx")
                index = doc.range.fields[0].as_field_index()

                if sort_entries_using_yomi:
                    self.assertTrue(index.use_yomi)
                    self.assertEqual(" INDEX  \\y", index.get_field_code())
                    self.assertEqual("愛子, 2\r" +
                                     "明美, 3\r" +
                                     "恵美, 4\r" +
                                     "愛美, 5\r", index.result)
                else:
                    self.assertFalse(index.use_yomi)
                    self.assertEqual(" INDEX ", index.get_field_code())
                    self.assertEqual("恵美, 4\r" +
                                     "愛子, 2\r" +
                                     "愛美, 5\r" +
                                     "明美, 3\r", index.result)

                index_entry = doc.range.fields[1].as_field_xe()

                self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " XE  愛子 \\y あ", "", index_entry)
                self.assertEqual("愛子", index_entry.text)
                self.assertEqual("あ", index_entry.yomi)

                index_entry = doc.range.fields[2].as_field_xe()

                self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " XE  明美 \\y あ", "", index_entry)
                self.assertEqual("明美", index_entry.text)
                self.assertEqual("あ", index_entry.yomi)

                index_entry = doc.range.fields[3].as_field_xe()

                self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " XE  恵美 \\y え", "", index_entry)
                self.assertEqual("恵美", index_entry.text)
                self.assertEqual("え", index_entry.yomi)

                index_entry = doc.range.fields[4].as_field_xe()

                self.verify_field(aw.fields.FieldType.FIELD_INDEX_ENTRY, " XE  愛美 \\y え", "", index_entry)
                self.assertEqual("愛美", index_entry.text)
                self.assertEqual("え", index_entry.yomi)

    def test_field_barcode(self):

        #ExStart
        #ExFor:FieldBarcode
        #ExFor:FieldBarcode.facing_identification_mark
        #ExFor:FieldBarcode.is_bookmark
        #ExFor:FieldBarcode.is_uspostal_address
        #ExFor:FieldBarcode.postal_address
        #ExSummary:Shows how to use the BARCODE field to display U.S. ZIP codes in the form of a barcode.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln()

        # Below are two ways of using BARCODE fields to display custom values as barcodes.
        # 1 -  Store the value that the barcode will display in the postal_address property:
        field = builder.insert_field(aw.fields.FieldType.FIELD_BARCODE, True).as_field_barcode()

        # This value needs to be a valid ZIP code.
        field.postal_address = "96801"
        field.is_uspostal_address = True
        field.facing_identification_mark = "C"

        self.assertEqual(" BARCODE  96801 \\u \\f C", field.get_field_code())

        builder.insert_break(aw.BreakType.LINE_BREAK)

        # 2 -  Reference a bookmark that stores the value that this barcode will display:
        field = builder.insert_field(aw.fields.FieldType.FIELD_BARCODE, True).as_field_barcode()
        field.postal_address = "BarcodeBookmark"
        field.is_bookmark = True

        self.assertEqual(" BARCODE  BarcodeBookmark \\b", field.get_field_code())

        # The bookmark that the BARCODE field references in its "postal_address" property
        # need to contain nothing besides the valid ZIP code.
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.start_bookmark("BarcodeBookmark")
        builder.writeln("968877")
        builder.end_bookmark("BarcodeBookmark")

        doc.save(ARTIFACTS_DIR + "Field.field_barcode.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_barcode.docx")

        self.assertEqual(0, doc.get_child_nodes(aw.NodeType.SHAPE, True).count)

        field = doc.range.fields[0].as_field_barcode()

        self.verify_field(aw.fields.FieldType.FIELD_BARCODE, " BARCODE  96801 \\u \\f C", "", field)
        self.assertEqual("C", field.facing_identification_mark)
        self.assertEqual("96801", field.postal_address)
        self.assertTrue(field.is_uspostal_address)

        field = doc.range.fields[1].as_field_barcode()

        self.verify_field(aw.fields.FieldType.FIELD_BARCODE, " BARCODE  BarcodeBookmark \\b", "", field)
        self.assertEqual("BarcodeBookmark", field.postal_address)
        self.assertTrue(field.is_bookmark)

    def test_field_display_barcode(self):

        #ExStart
        #ExFor:FieldDisplayBarcode
        #ExFor:FieldDisplayBarcode.add_start_stop_char
        #ExFor:FieldDisplayBarcode.background_color
        #ExFor:FieldDisplayBarcode.barcode_type
        #ExFor:FieldDisplayBarcode.barcode_value
        #ExFor:FieldDisplayBarcode.case_code_style
        #ExFor:FieldDisplayBarcode.display_text
        #ExFor:FieldDisplayBarcode.error_correction_level
        #ExFor:FieldDisplayBarcode.fix_check_digit
        #ExFor:FieldDisplayBarcode.foreground_color
        #ExFor:FieldDisplayBarcode.pos_code_style
        #ExFor:FieldDisplayBarcode.scaling_factor
        #ExFor:FieldDisplayBarcode.symbol_height
        #ExFor:FieldDisplayBarcode.symbol_rotation
        #ExSummary:Shows how to insert a DISPLAYBARCODE field, and set its properties.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        field = builder.insert_field(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, True).as_field_display_barcode()

        # Below are four types of barcodes, decorated in various ways, that the DISPLAYBARCODE field can display.
        # 1 -  QR code with custom colors:
        field.barcode_type = "QR"
        field.barcode_value = "ABC123"
        field.background_color = "0xF8BD69"
        field.foreground_color = "0xB5413B"
        field.error_correction_level = "3"
        field.scaling_factor = "250"
        field.symbol_height = "1000"
        field.symbol_rotation = "0"

        self.assertEqual(" DISPLAYBARCODE  ABC123 QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0", field.get_field_code())
        builder.writeln()

        # 2 -  EAN13 barcode, with the digits displayed below the bars:
        field = builder.insert_field(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, True).as_field_display_barcode()
        field.barcode_type = "EAN13"
        field.barcode_value = "501234567890"
        field.display_text = True
        field.pos_code_style = "CASE"
        field.fix_check_digit = True

        self.assertEqual(" DISPLAYBARCODE  501234567890 EAN13 \\t \\p CASE \\x", field.get_field_code())
        builder.writeln()

        # 3 -  CODE39 barcode:
        field = builder.insert_field(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, True).as_field_display_barcode()
        field.barcode_type = "CODE39"
        field.barcode_value = "12345ABCDE"
        field.add_start_stop_char = True

        self.assertEqual(" DISPLAYBARCODE  12345ABCDE CODE39 \\d", field.get_field_code())
        builder.writeln()

        # 4 -  ITF4 barcode, with a specified case code:
        field = builder.insert_field(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, True).as_field_display_barcode()
        field.barcode_type = "ITF14"
        field.barcode_value = "09312345678907"
        field.case_code_style = "STD"

        self.assertEqual(" DISPLAYBARCODE  09312345678907 ITF14 \\c STD", field.get_field_code())

        doc.save(ARTIFACTS_DIR + "Field.field_display_barcode.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_display_barcode.docx")

        self.assertEqual(0, doc.get_child_nodes(aw.NodeType.SHAPE, True).count)

        field = doc.range.fields[0].as_field_display_barcode()

        self.verify_field(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, " DISPLAYBARCODE  ABC123 QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0", "", field)
        self.assertEqual("QR", field.barcode_type)
        self.assertEqual("ABC123", field.barcode_value)
        self.assertEqual("0xF8BD69", field.background_color)
        self.assertEqual("0xB5413B", field.foreground_color)
        self.assertEqual("3", field.error_correction_level)
        self.assertEqual("250", field.scaling_factor)
        self.assertEqual("1000", field.symbol_height)
        self.assertEqual("0", field.symbol_rotation)

        field = doc.range.fields[1].as_field_display_barcode()

        self.verify_field(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, " DISPLAYBARCODE  501234567890 EAN13 \\t \\p CASE \\x", "", field)
        self.assertEqual("EAN13", field.barcode_type)
        self.assertEqual("501234567890", field.barcode_value)
        self.assertTrue(field.display_text)
        self.assertEqual("CASE", field.pos_code_style)
        self.assertTrue(field.fix_check_digit)

        field = doc.range.fields[2].as_field_display_barcode()

        self.verify_field(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, " DISPLAYBARCODE  12345ABCDE CODE39 \\d", "", field)
        self.assertEqual("CODE39", field.barcode_type)
        self.assertEqual("12345ABCDE", field.barcode_value)
        self.assertTrue(field.add_start_stop_char)

        field = doc.range.fields[3].as_field_display_barcode()

        self.verify_field(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, " DISPLAYBARCODE  09312345678907 ITF14 \\c STD", "", field)
        self.assertEqual("ITF14", field.barcode_type)
        self.assertEqual("09312345678907", field.barcode_value)
        self.assertEqual("STD", field.case_code_style)

    #def test_field_merge_barcode__q_r(self):

    #    #ExStart
    #    #ExFor:FieldDisplayBarcode
    #    #ExFor:FieldMergeBarcode
    #    #ExFor:FieldMergeBarcode.background_color
    #    #ExFor:FieldMergeBarcode.barcode_type
    #    #ExFor:FieldMergeBarcode.barcode_value
    #    #ExFor:FieldMergeBarcode.error_correction_level
    #    #ExFor:FieldMergeBarcode.foreground_color
    #    #ExFor:FieldMergeBarcode.scaling_factor
    #    #ExFor:FieldMergeBarcode.symbol_height
    #    #ExFor:FieldMergeBarcode.symbol_rotation
    #    #ExSummary:Shows how to perform a mail merge on QR barcodes.
    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    # Insert a MERGEBARCODE field, which will accept values from a data source during a mail merge.
    #    # This field will convert all values in a merge data source's "MyQRCode" column into QR codes.
    #    field = builder.insert_field(aw.fields.FieldType.FIELD_MERGE_BARCODE, True).as_field_merge_barcode()
    #    field.barcode_type = "QR"
    #    field.barcode_value = "MyQRCode"

    #    # Apply custom colors and scaling.
    #    field.background_color = "0xF8BD69"
    #    field.foreground_color = "0xB5413B"
    #    field.error_correction_level = "3"
    #    field.scaling_factor = "250"
    #    field.symbol_height = "1000"
    #    field.symbol_rotation = "0"

    #    self.assertEqual(aw.fields.FieldType.FIELD_MERGE_BARCODE, field.type)
    #    self.assertEqual(" MERGEBARCODE  MyQRCode QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0",
    #        field.get_field_code())
    #    builder.writeln()

    #    # Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue.
    #    # The mail merge will create a new page for each row. Each page will contain a DISPLAYBARCODE field,
    #    # which will display a QR code with the value from the merged row.
    #    table = DataTable("Barcodes")
    #    table.columns.add("MyQRCode")
    #    table.rows.add(["ABC123"])
    #    table.rows.add(["DEF456"])

    #    doc.mail_merge.execute(table)

    #    self.assertEqual(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, doc.range.fields[0].type)
    #    self.assertEqual("DISPLAYBARCODE \"ABC123\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B",
    #        doc.range.fields[0].get_field_code())
    #    self.assertEqual(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, doc.range.fields[1].type)
    #    self.assertEqual("DISPLAYBARCODE \"DEF456\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B",
    #        doc.range.fields[1].get_field_code())

    #    doc.save(ARTIFACTS_DIR + "Field.mergebarcode.qr.docx")
    #    #ExEnd

    #    doc = aw.Document(ARTIFACTS_DIR + "Field.mergebarcode.qr.docx")

    #    self.assertEqual(0, len([f for f in doc.range.fields if f.type == aw.fields.FieldType.FIELD_MERGE_BARCODE]))

    #    barcode = doc.range.fields[0].as_field_display_barcode()

    #    self.verify_field(aw.fields.FieldType.FIELD_DISPLAY_BARCODE,
    #        "DISPLAYBARCODE \"ABC123\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B", "", barcode)
    #    self.assertEqual("ABC123", barcode.barcode_value)
    #    self.assertEqual("QR", barcode.barcode_type)

    #    barcode = doc.range.fields[1].as_field_display_barcode()

    #    self.verify_field(aw.fields.FieldType.FIELD_DISPLAY_BARCODE,
    #        "DISPLAYBARCODE \"DEF456\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B", "", barcode)
    #    self.assertEqual("DEF456", barcode.barcode_value)
    #    self.assertEqual("QR", barcode.barcode_type)

    #def test_field_merge_barcode_ean13(self):

    #    #ExStart
    #    #ExFor:FieldMergeBarcode
    #    #ExFor:FieldMergeBarcode.barcode_type
    #    #ExFor:FieldMergeBarcode.barcode_value
    #    #ExFor:FieldMergeBarcode.display_text
    #    #ExFor:FieldMergeBarcode.fix_check_digit
    #    #ExFor:FieldMergeBarcode.pos_code_style
    #    #ExSummary:Shows how to perform a mail merge on EAN13 barcodes.
    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    # Insert a MERGEBARCODE field, which will accept values from a data source during a mail merge.
    #    # This field will convert all values in a merge data source's "MyEAN13Barcode" column into EAN13 barcodes.
    #    field = builder.insert_field(aw.fields.FieldType.FIELD_MERGE_BARCODE, True).as_field_merge_barcode()
    #    field.barcode_type = "EAN13"
    #    field.barcode_value = "MyEAN13Barcode"

    #    # Display the numeric value of the barcode underneath the bars.
    #    field.display_text = True
    #    field.pos_code_style = "CASE"
    #    field.fix_check_digit = True

    #    self.assertEqual(aw.fields.FieldType.FIELD_MERGE_BARCODE, field.type)
    #    self.assertEqual(" MERGEBARCODE  MyEAN13Barcode EAN13 \\t \\p CASE \\x", field.get_field_code())
    #    builder.writeln()

    #    # Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue.
    #    # The mail merge will create a new page for each row. Each page will contain a DISPLAYBARCODE field,
    #    # which will display an EAN13 barcode with the value from the merged row.
    #    table = DataTable("Barcodes")
    #    table.columns.add("MyEAN13Barcode")
    #    table.rows.add(["501234567890"])
    #    table.rows.add(["123456789012"])

    #    doc.mail_merge.execute(table)

    #    self.assertEqual(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, doc.range.fields[0].type)
    #    self.assertEqual("DISPLAYBARCODE \"501234567890\" EAN13 \\t \\p CASE \\x",
    #        doc.range.fields[0].get_field_code())
    #    self.assertEqual(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, doc.range.fields[1].type)
    #    self.assertEqual("DISPLAYBARCODE \"123456789012\" EAN13 \\t \\p CASE \\x",
    #        doc.range.fields[1].get_field_code())

    #    doc.save(ARTIFACTS_DIR + "Field.mergebarcode.ean13.docx")
    #    #ExEnd

    #    doc = aw.Document(ARTIFACTS_DIR + "Field.mergebarcode.ean13.docx")

    #    self.assertEqual(0, len([f for f in doc.range.fields if f.type == aw.fields.FieldType.FIELD_MERGE_BARCODE]))

    #    barcode = doc.range.fields[0].as_field_display_barcode()

    #    self.verify_field(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, "DISPLAYBARCODE \"501234567890\" EAN13 \\t \\p CASE \\x", "", barcode)
    #    self.assertEqual("501234567890", barcode.barcode_value)
    #    self.assertEqual("EAN13", barcode.barcode_type)

    #    barcode = doc.range.fields[1].as_field_display_barcode()

    #    self.verify_field(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, "DISPLAYBARCODE \"123456789012\" EAN13 \\t \\p CASE \\x", "", barcode)
    #    self.assertEqual("123456789012", barcode.barcode_value)
    #    self.assertEqual("EAN13", barcode.barcode_type)

    #def test_field_merge_barcode_code39(self):

    #    #ExStart
    #    #ExFor:FieldMergeBarcode
    #    #ExFor:FieldMergeBarcode.add_start_stop_char
    #    #ExFor:FieldMergeBarcode.barcode_type
    #    #ExSummary:Shows how to perform a mail merge on CODE39 barcodes.
    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    # Insert a MERGEBARCODE field, which will accept values from a data source during a mail merge.
    #    # This field will convert all values in a merge data source's "MyCODE39Barcode" column into CODE39 barcodes.
    #    field = builder.insert_field(aw.fields.FieldType.FIELD_MERGE_BARCODE, True).as_field_merge_barcode()
    #    field.barcode_type = "CODE39"
    #    field.barcode_value = "MyCODE39Barcode"

    #    # Edit its appearance to display start/stop characters.
    #    field.add_start_stop_char = True

    #    self.assertEqual(aw.fields.FieldType.FIELD_MERGE_BARCODE, field.type)
    #    self.assertEqual(" MERGEBARCODE  MyCODE39Barcode CODE39 \\d", field.get_field_code())
    #    builder.writeln()

    #    # Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue.
    #    # The mail merge will create a new page for each row. Each page will contain a DISPLAYBARCODE field,
    #    # which will display a CODE39 barcode with the value from the merged row.
    #    table = DataTable("Barcodes")
    #    table.columns.add("MyCODE39Barcode")
    #    table.rows.add(["12345ABCDE"])
    #    table.rows.add(["67890FGHIJ"])

    #    doc.mail_merge.execute(table)

    #    self.assertEqual(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, doc.range.fields[0].type)
    #    self.assertEqual("DISPLAYBARCODE \"12345ABCDE\" CODE39 \\d",
    #        doc.range.fields[0].get_field_code())
    #    self.assertEqual(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, doc.range.fields[1].type)
    #    self.assertEqual("DISPLAYBARCODE \"67890FGHIJ\" CODE39 \\d",
    #        doc.range.fields[1].get_field_code())

    #    doc.save(ARTIFACTS_DIR + "Field.mergebarcode.code39.docx")
    #    #ExEnd

    #    doc = aw.Document(ARTIFACTS_DIR + "Field.mergebarcode.code39.docx")

    #    self.assertEqual(0, len([f for f in doc.range.fields if f.type == aw.fields.FieldType.FIELD_MERGE_BARCODE]))

    #    barcode = doc.range.fields[0].as_field_display_barcode()

    #    self.verify_field(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, "DISPLAYBARCODE \"12345ABCDE\" CODE39 \\d", "", barcode)
    #    self.assertEqual("12345ABCDE", barcode.barcode_value)
    #    self.assertEqual("CODE39", barcode.barcode_type)

    #    barcode = doc.range.fields[1].as_field_display_barcode()

    #    self.verify_field(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, "DISPLAYBARCODE \"67890FGHIJ\" CODE39 \\d", "", barcode)
    #    self.assertEqual("67890FGHIJ", barcode.barcode_value)
    #    self.assertEqual("CODE39", barcode.barcode_type)

    #def test_field_merge_barcode_itf14(self):

    #    #ExStart
    #    #ExFor:FieldMergeBarcode
    #    #ExFor:FieldMergeBarcode.barcode_type
    #    #ExFor:FieldMergeBarcode.case_code_style
    #    #ExSummary:Shows how to perform a mail merge on ITF14 barcodes.
    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    # Insert a MERGEBARCODE field, which will accept values from a data source during a mail merge.
    #    # This field will convert all values in a merge data source's "MyITF14Barcode" column into ITF14 barcodes.
    #    field = builder.insert_field(aw.fields.FieldType.FIELD_MERGE_BARCODE, True).as_field_merge_barcode()
    #    field.barcode_type = "ITF14"
    #    field.barcode_value = "MyITF14Barcode"
    #    field.case_code_style = "STD"

    #    self.assertEqual(aw.fields.FieldType.FIELD_MERGE_BARCODE, field.type)
    #    self.assertEqual(" MERGEBARCODE  MyITF14Barcode ITF14 \\c STD", field.get_field_code())

    #    # Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue.
    #    # The mail merge will create a new page for each row. Each page will contain a DISPLAYBARCODE field,
    #    # which will display an ITF14 barcode with the value from the merged row.
    #    table = DataTable("Barcodes")
    #    table.columns.add("MyITF14Barcode")
    #    table.rows.add(["09312345678907"])
    #    table.rows.add(["1234567891234"])

    #    doc.mail_merge.execute(table)

    #    self.assertEqual(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, doc.range.fields[0].type)
    #    self.assertEqual("DISPLAYBARCODE \"09312345678907\" ITF14 \\c STD",
    #        doc.range.fields[0].get_field_code())
    #    self.assertEqual(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, doc.range.fields[1].type)
    #    self.assertEqual("DISPLAYBARCODE \"1234567891234\" ITF14 \\c STD",
    #        doc.range.fields[1].get_field_code())

    #    doc.save(ARTIFACTS_DIR + "Field.mergebarcode.itf14.docx")
    #    #ExEnd

    #    doc = aw.Document(ARTIFACTS_DIR + "Field.mergebarcode.itf14.docx")

    #    self.assertEqual(0, len([f for f in doc.range.fields if f.type == aw.fields.FieldType.FIELD_MERGE_BARCODE]))

    #    barcode = doc.range.fields[0].as_field_display_barcode()

    #    self.verify_field(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, "DISPLAYBARCODE \"09312345678907\" ITF14 \\c STD", "", barcode)
    #    self.assertEqual("09312345678907", barcode.barcode_value)
    #    self.assertEqual("ITF14", barcode.barcode_type)

    #    barcode = doc.range.fields[1].as_field_display_barcode()

    #    self.verify_field(aw.fields.FieldType.FIELD_DISPLAY_BARCODE, "DISPLAYBARCODE \"1234567891234\" ITF14 \\c STD", "", barcode)
    #    self.assertEqual("1234567891234", barcode.barcode_value)
    #    self.assertEqual("ITF14", barcode.barcode_type)

    #ExStart
    #ExFor:FieldLink
    #ExFor:FieldLink.auto_update
    #ExFor:FieldLink.format_update_type
    #ExFor:FieldLink.insert_as_bitmap
    #ExFor:FieldLink.insert_as_html
    #ExFor:FieldLink.insert_as_picture
    #ExFor:FieldLink.insert_as_rtf
    #ExFor:FieldLink.insert_as_text
    #ExFor:FieldLink.insert_as_unicode
    #ExFor:FieldLink.is_linked
    #ExFor:FieldLink.prog_id
    #ExFor:FieldLink.source_full_name
    #ExFor:FieldLink.source_item
    #ExFor:FieldDde
    #ExFor:FieldDde.auto_update
    #ExFor:FieldDde.insert_as_bitmap
    #ExFor:FieldDde.insert_as_html
    #ExFor:FieldDde.insert_as_picture
    #ExFor:FieldDde.insert_as_rtf
    #ExFor:FieldDde.insert_as_text
    #ExFor:FieldDde.insert_as_unicode
    #ExFor:FieldDde.is_linked
    #ExFor:FieldDde.prog_id
    #ExFor:FieldDde.source_full_name
    #ExFor:FieldDde.source_item
    #ExFor:FieldDdeAuto
    #ExFor:FieldDdeAuto.insert_as_bitmap
    #ExFor:FieldDdeAuto.insert_as_html
    #ExFor:FieldDdeAuto.insert_as_picture
    #ExFor:FieldDdeAuto.insert_as_rtf
    #ExFor:FieldDdeAuto.insert_as_text
    #ExFor:FieldDdeAuto.insert_as_unicode
    #ExFor:FieldDdeAuto.is_linked
    #ExFor:FieldDdeAuto.prog_id
    #ExFor:FieldDdeAuto.source_full_name
    #ExFor:FieldDdeAuto.source_item
    #ExSummary:Shows how to use various field types to link to other documents in the local file system, and display their contents.
    @unittest.skip("WORDSNET-16226") #ExSkip
    def test_field_linked_objects_as_text(self):

        for insert_linked_object_as in (ExField.InsertLinkedObjectAs.TEXT,
                                        ExField.InsertLinkedObjectAs.UNICODE,
                                        ExField.InsertLinkedObjectAs.HTML,
                                        ExField.InsertLinkedObjectAs.RTF):
            with self.subTest(insert_linked_object_as=insert_linked_object_as):
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Below are three types of fields we can use to display contents from a linked document in the form of text.
                # 1 -  A LINK field:
                builder.writeln("FieldLink:\n")
                ExField.insert_field_link(builder, insert_linked_object_as, "Word.document.8", MY_DIR + "Document.docx", None, True)

                # 2 -  A DDE field:
                builder.writeln("FieldDde:\n")
                ExField.insert_field_fde(builder, insert_linked_object_as, "Excel.Sheet", MY_DIR + "Spreadsheet.xlsx",
                    "Sheet1!R1C1", True, True)

                # 3 -  A DDEAUTO field:
                builder.writeln("FieldDdeAuto:\n")
                ExField.insert_field_fde_auto(builder, insert_linked_object_as, "Excel.Sheet", MY_DIR + "Spreadsheet.xlsx",
                    "Sheet1!R1C1", True)

                doc.update_fields()
                doc.save(ARTIFACTS_DIR + "Field.field_linked_objects_as_text.docx")

    @unittest.skip("WORDSNET-16226") #ExSkip
    def test_field_linked_objects_as_image(self):

        for insert_linked_object_as in (ExField.InsertLinkedObjectAs.PICTURE,
                                        ExField.InsertLinkedObjectAs.BITMAP):
            with self.subTest(insert_linked_object_as=insert_linked_object_as):
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Below are three types of fields we can use to display contents from a linked document in the form of an image.
                # 1 -  A LINK field:
                builder.writeln("FieldLink:\n")
                ExField.insert_field_link(builder, insert_linked_object_as, "Excel.Sheet", MY_DIR + "MySpreadsheet.xlsx",
                    "Sheet1!R2C2", True)

                # 2 -  A DDE field:
                builder.writeln("FieldDde:\n")
                ExField.insert_field_fde(builder, insert_linked_object_as, "Excel.Sheet", MY_DIR + "Spreadsheet.xlsx",
                    "Sheet1!R1C1", True, True)

                # 3 -  A DDEAUTO field:
                builder.writeln("FieldDdeAuto:\n")
                ExField.insert_field_fde_auto(builder, insert_linked_object_as, "Excel.Sheet", MY_DIR + "Spreadsheet.xlsx",
                    "Sheet1!R1C1", True)

                doc.update_fields()
                doc.save(ARTIFACTS_DIR + "Field.field_linked_objects_as_image.docx")

    @staticmethod
    def insert_field_link(builder: aw.DocumentBuilder, insert_linked_object_as: 'ExField.InsertLinkedObjectAs',
        prog_id: str, source_full_name: str, source_item: str, should_auto_update: bool):
        """ExField.InsertLinkedObjectAs.BITMAP"""

        field = builder.insert_field(aw.fields.FieldType.FIELD_LINK, True).as_field_link()

        if insert_linked_object_as == ExField.InsertLinkedObjectAs.TEXT:
            field.insert_as_text = True

        elif insert_linked_object_as == ExField.InsertLinkedObjectAs.UNICODE:
            field.insert_as_unicode = True

        elif insert_linked_object_as == ExField.InsertLinkedObjectAs.HTML:
            field.insert_as_html = True

        elif insert_linked_object_as == ExField.InsertLinkedObjectAs.RTF:
            field.insert_as_rtf = True

        elif insert_linked_object_as == ExField.InsertLinkedObjectAs.PICTURE:
            field.insert_as_picture = True

        elif insert_linked_object_as == ExField.InsertLinkedObjectAs.BITMAP:
            field.insert_as_bitmap = True

        field.auto_update = should_auto_update
        field.prog_id = prog_id
        field.source_full_name = source_full_name
        field.source_item = source_item

        builder.writeln("\n")

    @staticmethod
    def insert_field_fde(builder: aw.DocumentBuilder, insert_linked_object_as: 'ExField.InsertLinkedObjectAs', prog_id: str,
        source_full_name: str, source_item: str, is_linked: bool, should_auto_update: bool):
        """Use a document builder to insert a DDE field, and set its properties according to parameters."""

        field = builder.insert_field(aw.fields.FieldType.FIELD_DDE, True).as_field_dde()

        if insert_linked_object_as == ExField.InsertLinkedObjectAs.TEXT:
            field.insert_as_text = True

        elif insert_linked_object_as == ExField.InsertLinkedObjectAs.UNICODE:
            field.insert_as_unicode = True

        elif insert_linked_object_as == ExField.InsertLinkedObjectAs.HTML:
            field.insert_as_html = True

        elif insert_linked_object_as == ExField.InsertLinkedObjectAs.RTF:
            field.insert_as_rtf = True

        elif insert_linked_object_as == ExField.InsertLinkedObjectAs.PICTURE:
            field.insert_as_picture = True

        elif insert_linked_object_as == ExField.InsertLinkedObjectAs.BITMAP:
            field.insert_as_bitmap = True

        field.auto_update = should_auto_update
        field.prog_id = prog_id
        field.source_full_name = source_full_name
        field.source_item = source_item
        field.is_linked = is_linked

        builder.writeln("\n")

    @staticmethod
    def insert_field_fde_auto(builder: aw.DocumentBuilder, insert_linked_object_as: 'ExField.InsertLinkedObjectAs',
        prog_id: str, source_full_name: str, source_item: str, is_linked: bool):
        """Use a document builder to insert a DDEAUTO, field and set its properties according to parameters."""

        field = builder.insert_field(aw.fields.FieldType.FIELD_DDEAUTO, True).as_field_dde_auto()

        if insert_linked_object_as == ExField.InsertLinkedObjectAs.TEXT:
            field.insert_as_text = True

        elif insert_linked_object_as == ExField.InsertLinkedObjectAs.UNICODE:
            field.insert_as_unicode = True

        elif insert_linked_object_as == ExField.InsertLinkedObjectAs.HTML:
            field.insert_as_html = True

        elif insert_linked_object_as == ExField.InsertLinkedObjectAs.RTF:
            field.insert_as_rtf = True

        elif insert_linked_object_as == ExField.InsertLinkedObjectAs.PICTURE:
            field.insert_as_picture = True

        elif insert_linked_object_as == ExField.InsertLinkedObjectAs.BITMAP:
            field.insert_as_bitmap = True

        field.prog_id = prog_id
        field.source_full_name = source_full_name
        field.source_item = source_item
        field.is_linked = is_linked

    class InsertLinkedObjectAs(Enum):

        # LinkedObjectAsText
        TEXT = 1
        UNICODE = 2
        HTML = 3
        RTF = 4
        # LinkedObjectAsImage
        PICTURE = 5
        BITMAP = 6

    #ExEnd

    def test_field_user_address(self):

        #ExStart
        #ExFor:FieldUserAddress
        #ExFor:FieldUserAddress.user_address
        #ExSummary:Shows how to use the USERADDRESS field.
        doc = aw.Document()

        # Create a UserInformation object and set it as the source of user information for any fields that we create.
        user_information = aw.fields.UserInformation()
        user_information.address = "123 Main Street"
        doc.field_options.current_user = user_information

        # Create a USERADDRESS field to display the current user's address,
        # taken from the UserInformation object we created above.
        builder = aw.DocumentBuilder(doc)
        field_user_address = builder.insert_field(aw.fields.FieldType.FIELD_USER_ADDRESS, True).as_field_user_address()
        self.assertEqual(user_information.address, field_user_address.result) #ExSkip

        self.assertEqual(" USERADDRESS ", field_user_address.get_field_code())
        self.assertEqual("123 Main Street", field_user_address.result)

        # We can set this property to get our field to override the value currently stored in the UserInformation object.
        field_user_address.user_address = "456 North Road"
        field_user_address.update()

        self.assertEqual(" USERADDRESS  \"456 North Road\"", field_user_address.get_field_code())
        self.assertEqual("456 North Road", field_user_address.result)

        # This does not affect the value in the UserInformation object.
        self.assertEqual("123 Main Street", doc.field_options.current_user.address)

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_user_address.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_user_address.docx")

        field_user_address = doc.range.fields[0].as_field_user_address()

        self.verify_field(aw.fields.FieldType.FIELD_USER_ADDRESS, " USERADDRESS  \"456 North Road\"", "456 North Road", field_user_address)
        self.assertEqual("456 North Road", field_user_address.user_address)

    def test_field_user_initials(self):

        #ExStart
        #ExFor:FieldUserInitials
        #ExFor:FieldUserInitials.user_initials
        #ExSummary:Shows how to use the USERINITIALS field.
        doc = aw.Document()

        # Create a UserInformation object and set it as the source of user information for any fields that we create.
        user_information = aw.fields.UserInformation()
        user_information.initials = "J. D."
        doc.field_options.current_user = user_information

        # Create a USERINITIALS field to display the current user's initials,
        # taken from the UserInformation object we created above.
        builder = aw.DocumentBuilder(doc)
        field_user_initials = builder.insert_field(aw.fields.FieldType.FIELD_USER_INITIALS, True).as_field_user_initials()
        self.assertEqual(user_information.initials, field_user_initials.result)

        self.assertEqual(" USERINITIALS ", field_user_initials.get_field_code())
        self.assertEqual("J. D.", field_user_initials.result)

        # We can set this property to get our field to override the value currently stored in the UserInformation object.
        field_user_initials.user_initials = "J. C."
        field_user_initials.update()

        self.assertEqual(" USERINITIALS  \"J. C.\"", field_user_initials.get_field_code())
        self.assertEqual("J. C.", field_user_initials.result)

        # This does not affect the value in the UserInformation object.
        self.assertEqual("J. D.", doc.field_options.current_user.initials)

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_user_initials.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_user_initials.docx")

        field_user_initials = doc.range.fields[0].as_field_user_initials()

        self.verify_field(aw.fields.FieldType.FIELD_USER_INITIALS, " USERINITIALS  \"J. C.\"", "J. C.", field_user_initials)
        self.assertEqual("J. C.", field_user_initials.user_initials)

    def test_field_user_name(self):

        #ExStart
        #ExFor:FieldUserName
        #ExFor:FieldUserName.user_name
        #ExSummary:Shows how to use the USERNAME field.
        doc = aw.Document()

        # Create a UserInformation object and set it as the source of user information for any fields that we create.
        user_information = aw.fields.UserInformation()
        user_information.name = "John Doe"
        doc.field_options.current_user = user_information

        builder = aw.DocumentBuilder(doc)

        # Create a USERNAME field to display the current user's name,
        # taken from the UserInformation object we created above.
        field_user_name = builder.insert_field(aw.fields.FieldType.FIELD_USER_NAME, True).as_field_user_name()
        self.assertEqual(user_information.name, field_user_name.result)

        self.assertEqual(" USERNAME ", field_user_name.get_field_code())
        self.assertEqual("John Doe", field_user_name.result)

        # We can set this property to get our field to override the value currently stored in the UserInformation object.
        field_user_name.user_name = "Jane Doe"
        field_user_name.update()

        self.assertEqual(" USERNAME  \"Jane Doe\"", field_user_name.get_field_code())
        self.assertEqual("Jane Doe", field_user_name.result)

        # This does not affect the value in the UserInformation object.
        self.assertEqual("John Doe", doc.field_options.current_user.name)

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_user_name.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_user_name.docx")

        field_user_name = doc.range.fields[0].as_field_user_name()

        self.verify_field(aw.fields.FieldType.FIELD_USER_NAME, " USERNAME  \"Jane Doe\"", "Jane Doe", field_user_name)
        self.assertEqual("Jane Doe", field_user_name.user_name)

    @unittest.skip("WORDSNET-17657")
    def test_field_style_ref_paragraph_numbers(self):

        #ExStart
        #ExFor:FieldStyleRef
        #ExFor:FieldStyleRef.insert_paragraph_number
        #ExFor:FieldStyleRef.insert_paragraph_number_in_full_context
        #ExFor:FieldStyleRef.insert_paragraph_number_in_relative_context
        #ExFor:FieldStyleRef.insert_relative_position
        #ExFor:FieldStyleRef.search_from_bottom
        #ExFor:FieldStyleRef.style_name
        #ExFor:FieldStyleRef.suppress_non_delimiters
        #ExSummary:Shows how to use STYLEREF fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create a list based using a Microsoft Word list template.
        list = doc.lists.add(aw.lists.ListTemplate.NUMBER_DEFAULT)

        # This generated list will display "1.a )".
        # Space before the bracket is a non-delimiter character, which we can suppress.
        list.list_levels[0].number_format = "\x0000."
        list.list_levels[1].number_format = "\x0001 )"

        # Add text and apply paragraph styles that STYLEREF fields will reference.
        builder.list_format.list = list
        builder.list_format.list_indent()
        builder.paragraph_format.style = doc.styles.get_by_name("List Paragraph")
        builder.writeln("Item 1")
        builder.paragraph_format.style = doc.styles.get_by_name("Quote")
        builder.writeln("Item 2")
        builder.paragraph_format.style = doc.styles.get_by_name("List Paragraph")
        builder.writeln("Item 3")
        builder.list_format.remove_numbers()
        builder.paragraph_format.style = doc.styles.get_by_name("Normal")

        # Place a STYLEREF field in the header and display the first "List Paragraph"-styled text in the document.
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        field = builder.insert_field(aw.fields.FieldType.FIELD_STYLE_REF, True).as_field_style_ref()
        field.style_name = "List Paragraph"

        # Place a STYLEREF field in the footer, and have it display the last text.
        builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)
        field = builder.insert_field(aw.fields.FieldType.FIELD_STYLE_REF, True).as_field_style_ref()
        field.style_name = "List Paragraph"
        field.search_from_bottom = True

        builder.move_to_document_end()

        # We can also use STYLEREF fields to reference the list numbers of lists.
        builder.write("\nParagraph number: ")
        field = builder.insert_field(aw.fields.FieldType.FIELD_STYLE_REF, True).as_field_style_ref()
        field.style_name = "Quote"
        field.insert_paragraph_number = True

        builder.write("\nParagraph number, relative context: ")
        field = builder.insert_field(aw.fields.FieldType.FIELD_STYLE_REF, True).as_field_style_ref()
        field.style_name = "Quote"
        field.insert_paragraph_number_in_relative_context = True

        builder.write("\nParagraph number, full context: ")
        field = builder.insert_field(aw.fields.FieldType.FIELD_STYLE_REF, True).as_field_style_ref()
        field.style_name = "Quote"
        field.insert_paragraph_number_in_full_context = True

        builder.write("\nParagraph number, full context, non-delimiter chars suppressed: ")
        field = builder.insert_field(aw.fields.FieldType.FIELD_STYLE_REF, True).as_field_style_ref()
        field.style_name = "Quote"
        field.insert_paragraph_number_in_full_context = True
        field.suppress_non_delimiters = True

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_style_ref_paragraph_numbers.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_style_ref_paragraph_numbers.docx")

        field = doc.range.fields[0].as_field_style_ref()

        self.verify_field(aw.fields.FieldType.FIELD_STYLE_REF, " STYLEREF  \"List Paragraph\"", "Item 1", field)
        self.assertEqual("List Paragraph", field.style_name)

        field = doc.range.fields[1].as_field_style_ref()

        self.verify_field(aw.fields.FieldType.FIELD_STYLE_REF, " STYLEREF  \"List Paragraph\" \\l", "Item 3", field)
        self.assertEqual("List Paragraph", field.style_name)
        self.assertTrue(field.search_from_bottom)

        field = doc.range.fields[2].as_field_style_ref()

        self.verify_field(aw.fields.FieldType.FIELD_STYLE_REF, " STYLEREF  Quote \\n", "b )", field)
        self.assertEqual("Quote", field.style_name)
        self.assertTrue(field.insert_paragraph_number)

        field = doc.range.fields[3].as_field_style_ref()

        self.verify_field(aw.fields.FieldType.FIELD_STYLE_REF, " STYLEREF  Quote \\r", "b )", field)
        self.assertEqual("Quote", field.style_name)
        self.assertTrue(field.insert_paragraph_number_in_relative_context)

        field = doc.range.fields[4].as_field_style_ref()

        self.verify_field(aw.fields.FieldType.FIELD_STYLE_REF, " STYLEREF  Quote \\w", "1.b )", field)
        self.assertEqual("Quote", field.style_name)
        self.assertTrue(field.insert_paragraph_number_in_full_context)

        field = doc.range.fields[5].as_field_style_ref()

        self.verify_field(aw.fields.FieldType.FIELD_STYLE_REF, " STYLEREF  Quote \\w \\t", "1.b)", field)
        self.assertEqual("Quote", field.style_name)
        self.assertTrue(field.insert_paragraph_number_in_full_context)
        self.assertTrue(field.suppress_non_delimiters)

    def test_field_date(self):

        #ExStart
        #ExFor:FieldDate
        #ExFor:FieldDate.use_lunar_calendar
        #ExFor:FieldDate.use_saka_era_calendar
        #ExFor:FieldDate.use_um_al_qura_calendar
        #ExFor:FieldDate.use_last_format
        #ExSummary:Shows how to use DATE fields to display dates according to different kinds of calendars.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # If we want the text in the document always to display the correct date, we can use a DATE field.
        # Below are three types of cultural calendars that a DATE field can use to display a date.
        # 1 -  Islamic Lunar Calendar:
        field = builder.insert_field(aw.fields.FieldType.FIELD_DATE, True).as_field_date()
        field.use_lunar_calendar = True
        self.assertEqual(" DATE  \\h", field.get_field_code())
        builder.writeln()

        # 2 -  Umm al-Qura calendar:
        field = builder.insert_field(aw.fields.FieldType.FIELD_DATE, True).as_field_date()
        field.use_um_al_qura_calendar = True
        self.assertEqual(" DATE  \\u", field.get_field_code())
        builder.writeln()

        # 3 -  Indian National Calendar:
        field = builder.insert_field(aw.fields.FieldType.FIELD_DATE, True).as_field_date()
        field.use_saka_era_calendar = True
        self.assertEqual(" DATE  \\s", field.get_field_code())
        builder.writeln()

        # Insert a DATE field and set its calendar type to the one last used by the host application.
        # In Microsoft Word, the type will be the most recently used in the Insert -> Text -> Date and Time dialog box.
        field = builder.insert_field(aw.fields.FieldType.FIELD_DATE, True).as_field_date()
        field.use_last_format = True
        self.assertEqual(" DATE  \\l", field.get_field_code())
        builder.writeln()

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_date.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_date.docx")

        field = doc.range.fields[0].as_field_date()

        self.assertEqual(aw.fields.FieldType.FIELD_DATE, field.type)
        self.assertTrue(field.use_lunar_calendar)
        self.assertEqual(" DATE  \\h", field.get_field_code())
        self.assertRegex(doc.range.fields[0].result, r"\d{1,2}[/]\d{1,2}[/]\d{4}")

        field = doc.range.fields[1].as_field_date()

        self.verify_field(aw.fields.FieldType.FIELD_DATE, " DATE  \\u", datetime.now().strftime("%d/%m/%Y"), field)
        self.assertTrue(field.use_um_al_qura_calendar)

        field = doc.range.fields[2].as_field_date()

        self.verify_field(aw.fields.FieldType.FIELD_DATE, " DATE  \\s", datetime.now().strftime("%d/%m/%Y"), field)
        self.assertTrue(field.use_saka_era_calendar)

        field = doc.range.fields[3].as_field_date()

        self.verify_field(aw.fields.FieldType.FIELD_DATE, " DATE  \\l", datetime.now().strftime("%d/%m/%Y"), field)
        self.assertTrue(field.use_last_format)

    @unittest.skip("WORDSNET-17669")
    def test_field_create_date(self):

        #ExStart
        #ExFor:FieldCreateDate
        #ExFor:FieldCreateDate.use_lunar_calendar
        #ExFor:FieldCreateDate.use_saka_era_calendar
        #ExFor:FieldCreateDate.use_um_al_qura_calendar
        #ExSummary:Shows how to use the CREATEDATE field to display the creation date/time of the document.
        doc = aw.Document(MY_DIR + "Document.docx")
        builder = aw.DocumentBuilder(doc)
        builder.move_to_document_end()
        builder.writeln(" Date this document was created:")

        # We can use the CREATEDATE field to display the date and time of the creation of the document.
        # Below are three different calendar types according to which the CREATEDATE field can display the date/time.
        # 1 -  Islamic Lunar Calendar:
        builder.write("According to the Lunar Calendar - ")
        field = builder.insert_field(aw.fields.FieldType.FIELD_CREATE_DATE, True).as_field_create_date()
        field.use_lunar_calendar = True

        self.assertEqual(" CREATEDATE  \\h", field.get_field_code())

        # 2 -  Umm al-Qura calendar:
        builder.write("\nAccording to the Umm al-Qura Calendar - ")
        field = builder.insert_field(aw.fields.FieldType.FIELD_CREATE_DATE, True).as_field_create_date()
        field.use_um_al_qura_calendar = True

        self.assertEqual(" CREATEDATE  \\u", field.get_field_code())

        # 3 -  Indian National Calendar:
        builder.write("\nAccording to the Indian National Calendar - ")
        field = builder.insert_field(aw.fields.FieldType.FIELD_CREATE_DATE, True).as_field_create_date()
        field.use_saka_era_calendar = True

        self.assertEqual(" CREATEDATE  \\s", field.get_field_code())

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_create_date.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_create_date.docx")

        self.assertEqual(datetime(2017, 12, 5, 9, 56, 0), doc.built_in_document_properties.created_time)

        expected_date = doc.built_in_document_properties.created_time.add_hours(TimeZoneInfo.local.get_utc_offset(datetime.utcnow()).hours)
        field = doc.range.fields[0].as_field_create_date()
        um_al_qura_calendar = UmAlQuraCalendar()

        self.verify_field(aw.fields.FieldType.FIELD_CREATE_DATE, " CREATEDATE  \\h",
            f"{umAlQuraCalendar.get_month(expected_date)}/{umAlQuraCalendar.get_day_of_month(expected_date)}/{umAlQuraCalendar.get_year(expected_date)} " +
            expected_date.add_hours(1).to_string("hh:mm:ss tt"), field)
        self.assertEqual(aw.fields.FieldType.FIELD_CREATE_DATE, field.type)
        self.assertTrue(field.use_lunar_calendar)

        field = doc.range.fields[1].as_field_create_date()

        self.verify_field(aw.fields.FieldType.FIELD_CREATE_DATE, " CREATEDATE  \\u",
            f"{umAlQuraCalendar.get_month(expected_date)}/{umAlQuraCalendar.get_day_of_month(expected_date)}/{umAlQuraCalendar.get_year(expected_date)} " +
            expected_date.add_hours(1).to_string("hh:mm:ss tt"), field)
        self.assertEqual(aw.fields.FieldType.FIELD_CREATE_DATE, field.type)
        self.assertTrue(field.use_um_al_qura_calendar)

    @unittest.skip("WORDSNET-17669")
    def test_field_save_date(self):

        #ExStart
        #ExFor:BuiltInDocumentProperties.last_saved_time
        #ExFor:FieldSaveDate
        #ExFor:FieldSaveDate.use_lunar_calendar
        #ExFor:FieldSaveDate.use_saka_era_calendar
        #ExFor:FieldSaveDate.use_um_al_qura_calendar
        #ExSummary:Shows how to use the SAVEDATE field to display the date/time of the document's most recent save operation performed using Microsoft Word.
        doc = aw.Document(MY_DIR + "Document.docx")
        builder = aw.DocumentBuilder(doc)
        builder.move_to_document_end()
        builder.writeln(" Date this document was last saved:")

        # We can use the SAVEDATE field to display the last save operation's date and time on the document.
        # The save operation that these fields refer to is the manual save in an application like Microsoft Word,
        # not the document's "save" method.
        # Below are three different calendar types according to which the SAVEDATE field can display the date/time.
        # 1 -  Islamic Lunar Calendar:
        builder.write("According to the Lunar Calendar - ")
        field = builder.insert_field(aw.fields.FieldType.FIELD_SAVE_DATE, True).as_field_save_date()
        field.use_lunar_calendar = True

        self.assertEqual(" SAVEDATE  \\h", field.get_field_code())

        # 2 -  Umm al-Qura calendar:
        builder.write("\nAccording to the Umm al-Qura calendar - ")
        field = builder.insert_field(aw.fields.FieldType.FIELD_SAVE_DATE, True).as_field_save_date()
        field.use_um_al_qura_calendar = True

        self.assertEqual(" SAVEDATE  \\u", field.get_field_code())

        # 3 -  Indian National calendar:
        builder.write("\nAccording to the Indian National calendar - ")
        field = builder.insert_field(aw.fields.FieldType.FIELD_SAVE_DATE, True).as_field_save_date()
        field.use_saka_era_calendar = True

        self.assertEqual(" SAVEDATE  \\s", field.get_field_code())

        # The SAVEDATE fields draw their date/time values from the "last_saved_time" built-in property.
        # The document's Save method will not update this value, but we can still update it manually.
        doc.built_in_document_properties.last_saved_time = datetime.now()

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_save_date.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_save_date.docx")

        print(doc.built_in_document_properties.last_saved_time)

        field = doc.range.fields[0].as_field_save_date()

        self.assertEqual(aw.fields.FieldType.FIELD_SAVE_DATE, field.type)
        self.assertTrue(field.use_lunar_calendar)
        self.assertEqual(" SAVEDATE  \\h", field.get_field_code())

        self.assertRegex(field.result, "\\d{1,2}[/]\\d{1,2}[/]\\d{4} \\d{1,2}:\\d{1,2}:\\d{1,2} [A,P]M")

        field = doc.range.fields[1].as_field_save_date()

        self.assertEqual(aw.fields.FieldType.FIELD_SAVE_DATE, field.type)
        self.assertTrue(field.use_um_al_qura_calendar)
        self.assertEqual(" SAVEDATE  \\u", field.get_field_code())
        self.assertRegex(field.result, "\\d{1,2}[/]\\d{1,2}[/]\\d{4} \\d{1,2}:\\d{1,2}:\\d{1,2} [A,P]M")

    def test_field_builder(self):

        #ExStart
        #ExFor:FieldBuilder
        #ExFor:FieldBuilder.add_argument(int)
        #ExFor:FieldBuilder.add_argument(FieldArgumentBuilder)
        #ExFor:FieldBuilder.add_argument(str)
        #ExFor:FieldBuilder.add_argument(float)
        #ExFor:FieldBuilder.add_argument(FieldBuilder)
        #ExFor:FieldBuilder.add_switch(str)
        #ExFor:FieldBuilder.add_switch(str,float)
        #ExFor:FieldBuilder.add_switch(str,int)
        #ExFor:FieldBuilder.add_switch(str,str)
        #ExFor:FieldBuilder.build_and_insert(Paragraph)
        #ExFor:FieldArgumentBuilder
        #ExFor:FieldArgumentBuilder.add_field(FieldBuilder)
        #ExFor:FieldArgumentBuilder.add_text(str)
        #ExFor:FieldArgumentBuilder.add_node(Inline)
        #ExSummary:Shows how to construct fields using a field builder, and then insert them into the document.
        doc = aw.Document()

        # Below are three examples of field construction done using a field builder.
        # 1 -  Single field:
        # Use a field builder to add a SYMBOL field which displays the ƒ (Florin) symbol.
        builder = aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_SYMBOL)
        builder.add_argument(402)
        builder.add_switch("\\f", "Arial")
        builder.add_switch("\\s", 25)
        builder.add_switch("\\u")
        field = builder.build_and_insert(doc.first_section.body.first_paragraph)

        self.assertEqual(" SYMBOL 402 \\f Arial \\s 25 \\u ", field.get_field_code())

        # 2 -  Nested field:
        # Use a field builder to create a formula field used as an inner field by another field builder.
        inner_formula_builder = aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_FORMULA)
        inner_formula_builder.add_argument(100)
        inner_formula_builder.add_argument("+")
        inner_formula_builder.add_argument(74)

        # Create another builder for another SYMBOL field, and insert the formula field
        # that we have created above into the SYMBOL field as its argument.
        builder = aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_SYMBOL)
        builder.add_argument(inner_formula_builder)
        field = builder.build_and_insert(doc.first_section.body.append_paragraph(""))

        # The outer SYMBOL field will use the formula field result, 174, as its argument,
        # which will make the field display the ® (Registered Sign) symbol since its character number is 174.
        self.assertEqual(" SYMBOL \u0013 = 100 + 74 \u0014\u0015 ", field.get_field_code())

        # 3 -  Multiple nested fields and arguments:
        # Now, we will use a builder to create an IF field, which displays one of two custom string values,
        # depending on the True/False value of its expression. To get a True/False value
        # that determines which string the IF field displays, the IF field will test two numeric expressions for equality.
        # We will provide the two expressions in the form of formula fields, which we will nest inside the IF field.
        left_expression = aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_FORMULA)
        left_expression.add_argument(2)
        left_expression.add_argument("+")
        left_expression.add_argument(3)

        right_expression = aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_FORMULA)
        right_expression.add_argument(2.5)
        right_expression.add_argument("*")
        right_expression.add_argument(5.2)

        # Next, we will build two field arguments, which will serve as the True/False output strings for the IF field.
        # These arguments will reuse the output values of our numeric expressions.
        true_output = aw.fields.FieldArgumentBuilder()
        true_output.add_text("True, both expressions amount to ")
        true_output.add_field(left_expression)

        false_output = aw.fields.FieldArgumentBuilder()
        false_output.add_node(aw.Run(doc, "False, "))
        false_output.add_field(left_expression)
        false_output.add_node(aw.Run(doc, " does not equal "))
        false_output.add_field(right_expression)

        # Finally, we will create one more field builder for the IF field and combine all of the expressions.
        builder = aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_IF)
        builder.add_argument(left_expression)
        builder.add_argument("=")
        builder.add_argument(right_expression)
        builder.add_argument(true_output)
        builder.add_argument(false_output)
        field = builder.build_and_insert(doc.first_section.body.append_paragraph(""))

        self.assertEqual(" IF \u0013 = 2 + 3 \u0014\u0015 = \u0013 = 2.5 * 5.2 \u0014\u0015 " +
                        "\"True, both expressions amount to \u0013 = 2 + 3 \u0014\u0015\" " +
                        "\"False, \u0013 = 2 + 3 \u0014\u0015 does not equal \u0013 = 2.5 * 5.2 \u0014\u0015\" ", field.get_field_code())

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_builder.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_builder.docx")

        field_symbol = doc.range.fields[0].as_field_symbol()

        self.verify_field(aw.fields.FieldType.FIELD_SYMBOL, " SYMBOL 402 \\f Arial \\s 25 \\u ", "", field_symbol)
        self.assertEqual("ƒ", field_symbol.display_result)

        field_symbol = doc.range.fields[1].as_field_symbol()

        self.verify_field(aw.fields.FieldType.FIELD_SYMBOL, " SYMBOL \u0013 = 100 + 74 \u0014174\u0015 ", "", field_symbol)
        self.assertEqual("®", field_symbol.display_result)

        self.verify_field(aw.fields.FieldType.FIELD_FORMULA, " = 100 + 74 ", "174", doc.range.fields[2])

        self.verify_field(aw.fields.FieldType.FIELD_IF,
            " IF \u0013 = 2 + 3 \u00145\u0015 = \u0013 = 2.5 * 5.2 \u001413\u0015 " +
            "\"True, both expressions amount to \u0013 = 2 + 3 \u0014\u0015\" " +
            "\"False, \u0013 = 2 + 3 \u00145\u0015 does not equal \u0013 = 2.5 * 5.2 \u001413\u0015\" ",
            "False, 5 does not equal 13", doc.range.fields[3])

        with self.assertRaises(Exception):
            self.fields_are_nested(doc.range.fields[2], doc.range.fields[3])

        self.verify_field(aw.fields.FieldType.FIELD_FORMULA, " = 2 + 3 ", "5", doc.range.fields[4])
        self.fields_are_nested(doc.range.fields[4], doc.range.fields[3])

        self.verify_field(aw.fields.FieldType.FIELD_FORMULA, " = 2.5 * 5.2 ", "13", doc.range.fields[5])
        self.fields_are_nested(doc.range.fields[5], doc.range.fields[3])

        self.verify_field(aw.fields.FieldType.FIELD_FORMULA, " = 2 + 3 ", "", doc.range.fields[6])
        self.fields_are_nested(doc.range.fields[6], doc.range.fields[3])

        self.verify_field(aw.fields.FieldType.FIELD_FORMULA, " = 2 + 3 ", "5", doc.range.fields[7])
        self.fields_are_nested(doc.range.fields[7], doc.range.fields[3])

        self.verify_field(aw.fields.FieldType.FIELD_FORMULA, " = 2.5 * 5.2 ", "13", doc.range.fields[8])
        self.fields_are_nested(doc.range.fields[8], doc.range.fields[3])

    def test_field_author(self):

        #ExStart
        #ExFor:FieldAuthor
        #ExFor:FieldAuthor.author_name
        #ExFor:FieldOptions.default_document_author
        #ExSummary:Shows how to use an AUTHOR field to display a document creator's name.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # AUTHOR fields source their results from the built-in document property called "author".
        # If we create and save a document in Microsoft Word,
        # it will have our username in that property.
        # However, if we create a document programmatically using Aspose.Words,
        # the "author" property, by default, will be an empty string.
        self.assertEqual("", doc.built_in_document_properties.author)

        # Set a backup author name for AUTHOR fields to use
        # if the "author" property contains an empty string.
        doc.field_options.default_document_author = "Joe Bloggs"

        builder.write("This document was created by ")
        field = builder.insert_field(aw.fields.FieldType.FIELD_AUTHOR, True).as_field_author()
        field.update()

        self.assertEqual(" AUTHOR ", field.get_field_code())
        self.assertEqual("Joe Bloggs", field.result)

        # Updating an AUTHOR field that contains a value
        # will apply that value to the "author" built-in property.
        self.assertEqual("Joe Bloggs", doc.built_in_document_properties.author)

        # Changing this property, then updating the AUTHOR field will apply this value to the field.
        doc.built_in_document_properties.author = "John Doe"
        field.update()

        self.assertEqual(" AUTHOR ", field.get_field_code())
        self.assertEqual("John Doe", field.result)

        # If we update an AUTHOR field after changing its "name" property,
        # then the field will display the new name and apply the new name to the built-in property.
        field.author_name = "Jane Doe"
        field.update()

        self.assertEqual(" AUTHOR  \"Jane Doe\"", field.get_field_code())
        self.assertEqual("Jane Doe", field.result)

        # AUTHOR fields do not affect the "default_document_author" property.
        self.assertEqual("Jane Doe", doc.built_in_document_properties.author)
        self.assertEqual("Joe Bloggs", doc.field_options.default_document_author)

        doc.save(ARTIFACTS_DIR + "Field.field_author.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_author.docx")

        self.assertIsNone(doc.field_options.default_document_author)
        self.assertEqual("Jane Doe", doc.built_in_document_properties.author)

        field = doc.range.fields[0].as_field_author()

        self.verify_field(aw.fields.FieldType.FIELD_AUTHOR, " AUTHOR  \"Jane Doe\"", "Jane Doe", field)
        self.assertEqual("Jane Doe", field.author_name)

    def test_field_doc_variable(self):

        #ExStart
        #ExFor:FieldDocProperty
        #ExFor:FieldDocVariable
        #ExFor:FieldDocVariable.variable_name
        #ExSummary:Shows how to use DOCPROPERTY fields to display document properties and variables.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Below are two ways of using DOCPROPERTY fields.
        # 1 -  Display a built-in property:
        # Set a custom value for the "category" built-in property, then insert a DOCPROPERTY field that references it.
        doc.built_in_document_properties.category = "My category"

        field_doc_property = builder.insert_field(" DOCPROPERTY Category ").as_field_doc_property()
        field_doc_property.update()

        self.assertEqual(" DOCPROPERTY Category ", field_doc_property.get_field_code())
        self.assertEqual("My category", field_doc_property.result)

        builder.insert_paragraph()

        # 2 -  Display a custom document variable:
        # Define a custom variable, then reference that variable with a DOCPROPERTY field.
        self.assertEqual(0, len(list(doc.variables)))
        doc.variables.add("My variable", "My variable's value")

        field_doc_variable = builder.insert_field(aw.fields.FieldType.FIELD_DOC_VARIABLE, True).as_field_doc_variable()
        field_doc_variable.variable_name = "My Variable"
        field_doc_variable.update()

        self.assertEqual(" DOCVARIABLE  \"My Variable\"", field_doc_variable.get_field_code())
        self.assertEqual("My variable's value", field_doc_variable.result)

        doc.save(ARTIFACTS_DIR + "Field.field_doc_variable.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_doc_variable.docx")

        self.assertEqual("My category", doc.built_in_document_properties.category)

        field_doc_property = doc.range.fields[0].as_field_doc_property()

        self.verify_field(aw.fields.FieldType.FIELD_DOC_PROPERTY, " DOCPROPERTY Category ", "My category", field_doc_property)

        field_doc_variable = doc.range.fields[1].as_field_doc_variable()

        self.verify_field(aw.fields.FieldType.FIELD_DOC_VARIABLE, " DOCVARIABLE  \"My Variable\"", "My variable's value", field_doc_variable)
        self.assertEqual("My Variable", field_doc_variable.variable_name)

    def test_field_subject(self):

        #ExStart
        #ExFor:FieldSubject
        #ExFor:FieldSubject.text
        #ExSummary:Shows how to use the SUBJECT field.
        doc = aw.Document()

        # Set a value for the document's "subject" built-in property.
        doc.built_in_document_properties.subject = "My subject"

        # Create a SUBJECT field to display the value of that built-in property.
        builder = aw.DocumentBuilder(doc)
        field = builder.insert_field(aw.fields.FieldType.FIELD_SUBJECT, True).as_field_subject()
        field.update()

        self.assertEqual(" SUBJECT ", field.get_field_code())
        self.assertEqual("My subject", field.result)

        # If we give the SUBJECT field's "text" property value and update it, the field will
        # overwrite the current value of the "subject" built-in property with the value of its Text property,
        # and then display the new value.
        field.text = "My new subject"
        field.update()

        self.assertEqual(" SUBJECT  \"My new subject\"", field.get_field_code())
        self.assertEqual("My new subject", field.result)

        self.assertEqual("My new subject", doc.built_in_document_properties.subject)

        doc.save(ARTIFACTS_DIR + "Field.field_subject.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_subject.docx")

        self.assertEqual("My new subject", doc.built_in_document_properties.subject)

        field = doc.range.fields[0].as_field_subject()

        self.verify_field(aw.fields.FieldType.FIELD_SUBJECT, " SUBJECT  \"My new subject\"", "My new subject", field)
        self.assertEqual("My new subject", field.text)

    def test_field_comments(self):

        #ExStart
        #ExFor:FieldComments
        #ExFor:FieldComments.text
        #ExSummary:Shows how to use the COMMENTS field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Set a value for the document's "comments" built-in property.
        doc.built_in_document_properties.comments = "My comment."

        # Create a COMMENTS field to display the value of that built-in property.
        field = builder.insert_field(aw.fields.FieldType.FIELD_COMMENTS, True).as_field_comments()
        field.update()

        self.assertEqual(" COMMENTS ", field.get_field_code())
        self.assertEqual("My comment.", field.result)

        # If we give the COMMENTS field's Text property value and update it, the field will
        # overwrite the current value of the "comments" built-in property with the value of its Text property,
        # and then display the new value.
        field.text = "My overriding comment."
        field.update()

        self.assertEqual(" COMMENTS  \"My overriding comment.\"", field.get_field_code())
        self.assertEqual("My overriding comment.", field.result)

        doc.save(ARTIFACTS_DIR + "Field.field_comments.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_comments.docx")

        self.assertEqual("My overriding comment.", doc.built_in_document_properties.comments)

        field = doc.range.fields[0].as_field_comments()

        self.verify_field(aw.fields.FieldType.FIELD_COMMENTS, " COMMENTS  \"My overriding comment.\"", "My overriding comment.", field)
        self.assertEqual("My overriding comment.", field.text)

    def test_field_file_size(self):

        #ExStart
        #ExFor:FieldFileSize
        #ExFor:FieldFileSize.is_in_kilobytes
        #ExFor:FieldFileSize.is_in_megabytes
        #ExSummary:Shows how to display the file size of a document with a FILESIZE field.
        doc = aw.Document(MY_DIR + "Document.docx")

        self.assertEqual(18105, doc.built_in_document_properties.bytes)

        builder = aw.DocumentBuilder(doc)
        builder.move_to_document_end()
        builder.insert_paragraph()

        # Below are three different units of measure
        # with which FILESIZE fields can display the document's file size.
        # 1 -  Bytes:
        field = builder.insert_field(aw.fields.FieldType.FIELD_FILE_SIZE, True).as_field_file_size()
        field.update()

        self.assertEqual(" FILESIZE ", field.get_field_code())
        self.assertEqual("18105", field.result)

        # 2 -  Kilobytes:
        builder.insert_paragraph()
        field = builder.insert_field(aw.fields.FieldType.FIELD_FILE_SIZE, True).as_field_file_size()
        field.is_in_kilobytes = True
        field.update()

        self.assertEqual(" FILESIZE  \\k", field.get_field_code())
        self.assertEqual("18", field.result)

        # 3 -  Megabytes:
        builder.insert_paragraph()
        field = builder.insert_field(aw.fields.FieldType.FIELD_FILE_SIZE, True).as_field_file_size()
        field.is_in_megabytes = True
        field.update()

        self.assertEqual(" FILESIZE  \\m", field.get_field_code())
        self.assertEqual("0", field.result)

        # To update the values of these fields while editing in Microsoft Word,
        # we must first save the changes, and then manually update these fields.
        doc.save(ARTIFACTS_DIR + "Field.field_file_size.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_file_size.docx")

        field = doc.range.fields[0].as_field_file_size()

        self.verify_field(aw.fields.FieldType.FIELD_FILE_SIZE, " FILESIZE ", "18105", field)

        # These fields will need to be updated to produce an accurate result.
        doc.update_fields()

        field = doc.range.fields[1].as_field_file_size()

        self.verify_field(aw.fields.FieldType.FIELD_FILE_SIZE, " FILESIZE  \\k", "13", field)
        self.assertTrue(field.is_in_kilobytes)

        field = doc.range.fields[2].as_field_file_size()

        self.verify_field(aw.fields.FieldType.FIELD_FILE_SIZE, " FILESIZE  \\m", "0", field)
        self.assertTrue(field.is_in_megabytes)

    def test_field_go_to_button(self):

        #ExStart
        #ExFor:FieldGoToButton
        #ExFor:FieldGoToButton.display_text
        #ExFor:FieldGoToButton.location
        #ExSummary:Shows to insert a GOTOBUTTON field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Add a GOTOBUTTON field. When we double-click this field in Microsoft Word,
        # it will take the text cursor to the bookmark whose name the Location property references.
        field = builder.insert_field(aw.fields.FieldType.FIELD_GO_TO_BUTTON, True).as_field_go_to_button()
        field.display_text = "My Button"
        field.location = "MyBookmark"

        self.assertEqual(" GOTOBUTTON  MyBookmark My Button", field.get_field_code())

        # Insert a valid bookmark for the field to reference.
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.start_bookmark(field.location)
        builder.writeln("Bookmark text contents.")
        builder.end_bookmark(field.location)

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_go_to_button.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_go_to_button.docx")
        field = doc.range.fields[0].as_field_go_to_button()

        self.verify_field(aw.fields.FieldType.FIELD_GO_TO_BUTTON, " GOTOBUTTON  MyBookmark My Button", "", field)
        self.assertEqual("My Button", field.display_text)
        self.assertEqual("MyBookmark", field.location)

    ##ExStart
    ##ExFor:FieldFillIn
    ##ExFor:FieldFillIn.default_response
    ##ExFor:FieldFillIn.prompt_once_on_mail_merge
    ##ExFor:FieldFillIn.prompt_text
    ##ExSummary:Shows how to use the FILLIN field to prompt the user for a response.
    #def test_field_fill_in(self):

    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    # Insert a FILLIN field. When we manually update this field in Microsoft Word,
    #    # it will prompt us to enter a response. The field will then display the response as text.
    #    field = builder.insert_field(aw.fields.FieldType.FIELD_FILL_IN, True).as_field_fill_in()
    #    field.prompt_text = "Please enter a response:"
    #    field.default_response = "A default response."

    #    # We can also use these fields to ask the user for a unique response for each page
    #    # created during a mail merge done using Microsoft Word.
    #    field.prompt_once_on_mail_merge = True

    #    self.assertEqual(" FILLIN  \"Please enter a response:\" \\d \"A default response.\" \\o", field.get_field_code())

    #    merge_field = builder.insert_field(aw.fields.FieldType.FIELD_MERGE_FIELD, True).as_field_merge_field()
    #    mergeField.field_name = "MergeField"

    #    # If we perform a mail merge programmatically, we can use a custom prompt respondent
    #    # to automatically edit responses for FILLIN fields that the mail merge encounters.
    #    doc.field_options.user_prompt_respondent = ExField.PromptRespondent()
    #    doc.mail_merge.execute(["MergeField"], [""])

    #    doc.update_fields()
    #    doc.save(ARTIFACTS_DIR + "Field.fillin.docx")
    #    self._test_field_fill_in(aw.Document(ARTIFACTS_DIR + "Field.fillin.docx")) #ExSKip

    #class PromptRespondent(aw.fields.IFieldUserPromptRespondent):
    #    """Prepends a line to the default response of every FILLIN field during a mail merge."""

    #    def respond(self, prompt_text: str, default_response: str) -> str:

    #        return "Response modified by PromptRespondent. " + default_response

    ##ExEnd

    def _test_field_fill_in(self, doc: aw.Document):

        doc = DocumentHelper.save_open(doc)

        self.assertEqual(1, doc.range.fields.count)

        field = doc.range.fields[0].as_field_fill_in()

        self.verify_field(aw.fields.FieldType.FIELD_FILL_IN, " FILLIN  \"Please enter a response:\" \\d \"A default response.\" \\o",
            "Response modified by PromptRespondent. A default response.", field)
        self.assertEqual("Please enter a response:", field.prompt_text)
        self.assertEqual("A default response.", field.default_response)
        self.assertTrue(field.prompt_once_on_mail_merge)

    def test_field_info(self):

        #ExStart
        #ExFor:FieldInfo
        #ExFor:FieldInfo.info_type
        #ExFor:FieldInfo.new_value
        #ExSummary:Shows how to work with INFO fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Set a value for the "comments" built-in property and then insert an INFO field to display that property's value.
        doc.built_in_document_properties.comments = "My comment"
        field = builder.insert_field(aw.fields.FieldType.FIELD_INFO, True).as_field_info()
        field.info_type = "Comments"
        field.update()

        self.assertEqual(" INFO  Comments", field.get_field_code())
        self.assertEqual("My comment", field.result)

        builder.writeln()

        # Setting a value for the field's "new_value" property and updating
        # the field will also overwrite the corresponding built-in property with the new value.
        field = builder.insert_field(aw.fields.FieldType.FIELD_INFO, True).as_field_info()
        field.info_type = "Comments"
        field.new_value = "New comment"
        field.update()

        self.assertEqual(" INFO  Comments \"New comment\"", field.get_field_code())
        self.assertEqual("New comment", field.result)
        self.assertEqual("New comment", doc.built_in_document_properties.comments)

        doc.save(ARTIFACTS_DIR + "Field.field_info.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_info.docx")

        self.assertEqual("New comment", doc.built_in_document_properties.comments)

        field = doc.range.fields[0].as_field_info()

        self.verify_field(aw.fields.FieldType.FIELD_INFO, " INFO  Comments", "My comment", field)
        self.assertEqual("Comments", field.info_type)

        field = doc.range.fields[1].as_field_info()

        self.verify_field(aw.fields.FieldType.FIELD_INFO, " INFO  Comments \"New comment\"", "New comment", field)
        self.assertEqual("Comments", field.info_type)
        self.assertEqual("New comment", field.new_value)

    def test_field_macro_button(self):

        #ExStart
        #ExFor:Document.has_macros
        #ExFor:FieldMacroButton
        #ExFor:FieldMacroButton.display_text
        #ExFor:FieldMacroButton.macro_name
        #ExSummary:Shows how to use MACROBUTTON fields to allow us to run a document's macros by clicking.
        doc = aw.Document(MY_DIR + "Macro.docm")
        builder = aw.DocumentBuilder(doc)

        self.assertTrue(doc.has_macros)

        # Insert a MACROBUTTON field, and reference one of the document's macros by name in the "macro_name" property.
        field = builder.insert_field(aw.fields.FieldType.FIELD_MACRO_BUTTON, True).as_field_macro_button()
        field.macro_name = "MyMacro"
        field.display_text = "Double click to run macro: " + field.macro_name

        self.assertEqual(" MACROBUTTON  MyMacro Double click to run macro: MyMacro", field.get_field_code())

        # Use the property to reference "ViewZoom200", a macro that ships with Microsoft Word.
        # We can find all other macros via View -> Macros (dropdown) -> View Macros.
        # In that menu, select "Word Commands" from the "Macros in:" drop down.
        # If our document contains a custom macro with the same name as a stock macro,
        # our macro will be the one that the MACROBUTTON field runs.
        builder.insert_paragraph()
        field = builder.insert_field(aw.fields.FieldType.FIELD_MACRO_BUTTON, True).as_field_macro_button()
        field.macro_name = "ViewZoom200"
        field.display_text = "Run " + field.macro_name

        self.assertEqual(" MACROBUTTON  ViewZoom200 Run ViewZoom200", field.get_field_code())

        # Save the document as a macro-enabled document type.
        doc.save(ARTIFACTS_DIR + "Field.field_macro_button.docm")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_macro_button.docm")

        field = doc.range.fields[0].as_field_macro_button()

        self.verify_field(aw.fields.FieldType.FIELD_MACRO_BUTTON, " MACROBUTTON  MyMacro Double click to run macro: MyMacro", "", field)
        self.assertEqual("MyMacro", field.macro_name)
        self.assertEqual("Double click to run macro: MyMacro", field.display_text)

        field = doc.range.fields[1].as_field_macro_button()

        self.verify_field(aw.fields.FieldType.FIELD_MACRO_BUTTON, " MACROBUTTON  ViewZoom200 Run ViewZoom200", "", field)
        self.assertEqual("ViewZoom200", field.macro_name)
        self.assertEqual("Run ViewZoom200", field.display_text)

    def test_field_keywords(self):

        #ExStart
        #ExFor:FieldKeywords
        #ExFor:FieldKeywords.text
        #ExSummary:Shows to insert a KEYWORDS field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Add some keywords, also referred to as "tags" in File Explorer.
        doc.built_in_document_properties.keywords = "Keyword1, Keyword2"

        # The KEYWORDS field displays the value of this property.
        field = builder.insert_field(aw.fields.FieldType.FIELD_KEYWORD, True).as_field_keywords()
        field.update()

        self.assertEqual(" KEYWORDS ", field.get_field_code())
        self.assertEqual("Keyword1, Keyword2", field.result)

        # Setting a value for the field's "text" property,
        # and then updating the field will also overwrite the corresponding built-in property with the new value.
        field.text = "OverridingKeyword"
        field.update()

        self.assertEqual(" KEYWORDS  OverridingKeyword", field.get_field_code())
        self.assertEqual("OverridingKeyword", field.result)
        self.assertEqual("OverridingKeyword", doc.built_in_document_properties.keywords)

        doc.save(ARTIFACTS_DIR + "Field.field_keywords.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_keywords.docx")

        self.assertEqual("OverridingKeyword", doc.built_in_document_properties.keywords)

        field = doc.range.fields[0].as_field_keywords()

        self.verify_field(aw.fields.FieldType.FIELD_KEYWORD, " KEYWORDS  OverridingKeyword", "OverridingKeyword", field)
        self.assertEqual("OverridingKeyword", field.text)

    def test_field_num(self):

        #ExStart
        #ExFor:FieldPage
        #ExFor:FieldNumChars
        #ExFor:FieldNumPages
        #ExFor:FieldNumWords
        #ExSummary:Shows how to use NUMCHARS, NUMWORDS, NUMPAGES and PAGE fields to track the size of our documents.
        doc = aw.Document(MY_DIR + "Paragraphs.docx")
        builder = aw.DocumentBuilder(doc)

        builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER

        # Below are three types of fields that we can use to track the size of our documents.
        # 1 -  Track the character count with a NUMCHARS field:
        field_num_chars = builder.insert_field(aw.fields.FieldType.FIELD_NUM_CHARS, True).as_field_num_chars()
        builder.writeln(" characters")

        # 2 -  Track the word count with a NUMWORDS field:
        field_num_words = builder.insert_field(aw.fields.FieldType.FIELD_NUM_WORDS, True).as_field_num_words()
        builder.writeln(" words")

        # 3 -  Use both PAGE and NUMPAGES fields to display what page the field is on,
        # and the total number of pages in the document:
        builder.paragraph_format.alignment = aw.ParagraphAlignment.RIGHT
        builder.write("Page ")
        field_page = builder.insert_field(aw.fields.FieldType.FIELD_PAGE, True).as_field_page()
        builder.write(" of ")
        field_num_pages = builder.insert_field(aw.fields.FieldType.FIELD_NUM_PAGES, True).as_field_num_pages()

        self.assertEqual(" NUMCHARS ", field_num_chars.get_field_code())
        self.assertEqual(" NUMWORDS ", field_num_words.get_field_code())
        self.assertEqual(" NUMPAGES ", field_num_pages.get_field_code())
        self.assertEqual(" PAGE ", field_page.get_field_code())

        # These fields will not maintain accurate values in real time
        # while we edit the document programmatically using Aspose.Words, or in Microsoft Word.
        # We need to update them every we need to see an up-to-date value.
        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_num.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_num.docx")

        self.verify_field(aw.fields.FieldType.FIELD_NUM_CHARS, " NUMCHARS ", "6009", doc.range.fields[0])
        self.verify_field(aw.fields.FieldType.FIELD_NUM_WORDS, " NUMWORDS ", "1054", doc.range.fields[1])

        self.verify_field(aw.fields.FieldType.FIELD_PAGE, " PAGE ", "6", doc.range.fields[2])
        self.verify_field(aw.fields.FieldType.FIELD_NUM_PAGES, " NUMPAGES ", "6", doc.range.fields[3])

    def test_field_print(self):

        #ExStart
        #ExFor:FieldPrint
        #ExFor:FieldPrint.post_script_group
        #ExFor:FieldPrint.printer_instructions
        #ExSummary:Shows to insert a PRINT field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("My paragraph")

        # The PRINT field can send instructions to the printer.
        field = builder.insert_field(aw.fields.FieldType.FIELD_PRINT, True).as_field_print()

        # Set the area for the printer to perform instructions over.
        # In this case, it will be the paragraph that contains our PRINT field.
        field.post_script_group = "para"

        # When we use a printer that supports PostScript to print our document,
        # this command will turn the entire area that we specified in "post_script_group" white.
        field.printer_instructions = "erasepage"

        self.assertEqual(" PRINT  erasepage \\p para", field.get_field_code())

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_print.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_print.docx")

        field = doc.range.fields[0].as_field_print()

        self.verify_field(aw.fields.FieldType.FIELD_PRINT, " PRINT  erasepage \\p para", "", field)
        self.assertEqual("para", field.post_script_group)
        self.assertEqual("erasepage", field.printer_instructions)

    def test_field_print_date(self):

        #ExStart
        #ExFor:FieldPrintDate
        #ExFor:FieldPrintDate.use_lunar_calendar
        #ExFor:FieldPrintDate.use_saka_era_calendar
        #ExFor:FieldPrintDate.use_um_al_qura_calendar
        #ExSummary:Shows read PRINTDATE fields.
        doc = aw.Document(MY_DIR + "Field sample - PRINTDATE.docx")

        # When a document is printed by a printer or printed as a PDF (but not exported to PDF),
        # PRINTDATE fields will display the print operation's date/time.
        # If no printing has taken place, these fields will display "0/0/0000".
        field = doc.range.fields[0].as_field_print_date()

        self.assertEqual("3/25/2020 12:00:00 AM", field.result)
        self.assertEqual(" PRINTDATE ", field.get_field_code())

        # Below are three different calendar types according to which the PRINTDATE field
        # can display the date and time of the last printing operation.
        # 1 -  Islamic Lunar Calendar:
        field = doc.range.fields[1].as_field_print_date()

        self.assertTrue(field.use_lunar_calendar)
        self.assertEqual("8/1/1441 12:00:00 AM", field.result)
        self.assertEqual(" PRINTDATE  \\h", field.get_field_code())

        field = doc.range.fields[2].as_field_print_date()

        # 2 -  Umm al-Qura calendar:
        self.assertTrue(field.use_um_al_qura_calendar)
        self.assertEqual("8/1/1441 12:00:00 AM", field.result)
        self.assertEqual(" PRINTDATE  \\u", field.get_field_code())

        field = doc.range.fields[3].as_field_print_date()

        # 3 -  Indian National Calendar:
        self.assertTrue(field.use_saka_era_calendar)
        self.assertEqual("1/5/1942 12:00:00 AM", field.result)
        self.assertEqual(" PRINTDATE  \\s", field.get_field_code())
        #ExEnd

    def test_field_quote(self):

        #ExStart
        #ExFor:FieldQuote
        #ExFor:FieldQuote.text
        #ExFor:Document.update_fields
        #ExSummary:Shows to use the QUOTE field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a QUOTE field, which will display the value of its Text property.
        field = builder.insert_field(aw.fields.FieldType.FIELD_QUOTE, True).as_field_quote()
        field.text = "\"Quoted text\""

        self.assertEqual(" QUOTE  \"\\\"Quoted text\\\"\"", field.get_field_code())

        # Insert a QUOTE field and nest a DATE field inside it.
        # DATE fields update their value to the current date every time we open the document using Microsoft Word.
        # Nesting the DATE field inside the QUOTE field like this will freeze its value
        # to the date when we created the document.
        builder.write("\nDocument creation date: ")
        field = builder.insert_field(aw.fields.FieldType.FIELD_QUOTE, True).as_field_quote()
        builder.move_to(field.separator)
        builder.insert_field(aw.fields.FieldType.FIELD_DATE, True)

        self.assertEqual(" QUOTE \u0013 DATE \u0014" + datetime.now().strftime("%d/%m/%Y") + "\u0015", field.get_field_code())

        # Update all the fields to display their correct results.
        doc.update_fields()

        self.assertEqual("\"Quoted text\"", doc.range.fields[0].result)

        doc.save(ARTIFACTS_DIR + "Field.field_quote.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_quote.docx")

        self.verify_field(aw.fields.FieldType.FIELD_QUOTE, " QUOTE  \"\\\"Quoted text\\\"\"", "\"Quoted text\"", doc.range.fields[0])

        self.verify_field(aw.fields.FieldType.FIELD_QUOTE, " QUOTE \u0013 DATE \u0014" + datetime.now().strftime("%d/%m/%Y") + "\u0015",
            datetime.now().strftime("%d/%m/%Y"), doc.range.fields[1])

    ##ExStart
    ##ExFor:FieldNext
    ##ExFor:FieldNextIf
    ##ExFor:FieldNextIf.comparison_operator
    ##ExFor:FieldNextIf.left_expression
    ##ExFor:FieldNextIf.right_expression
    ##ExSummary:Shows how to use NEXT/NEXTIF fields to merge multiple rows into one page during a mail merge.
    #def test_field_next(self):

    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    # Create a data source for our mail merge with 3 rows.
    #    # A mail merge that uses this table would normally create a 3-page document.
    #    table = DataTable("Employees")
    #    table.columns.add("Courtesy Title")
    #    table.columns.add("First Name")
    #    table.columns.add("Last Name")
    #    table.rows.add("Mr.", "John", "Doe")
    #    table.rows.add("Mrs.", "Jane", "Cardholder")
    #    table.rows.add("Mr.", "Joe", "Bloggs")

    #    ExField.insert_merge_fields(builder, "First row: ")

    #    # If we have multiple merge fields with the same FieldName,
    #    # they will receive data from the same row of the data source and display the same value after the merge.
    #    # A NEXT field tells the mail merge instantly to move down one row,
    #    # which means any MERGEFIELDs that follow the NEXT field will receive data from the next row.
    #    # Make sure never to try to skip to the next row while already on the last row.
    #    field_next = builder.insert_field(aw.fields.FieldType.FIELD_NEXT, True).as_field_next()

    #    self.assertEqual(" NEXT ", field_next.get_field_code())

    #    # After the merge, the data source values that these MERGEFIELDs accept
    #    # will end up on the same page as the MERGEFIELDs above.
    #    ExField.insert_merge_fields(builder, "Second row: ")

    #    # A NEXTIF field has the same function as a NEXT field,
    #    # but it skips to the next row only if a statement constructed by the following 3 properties is True.
    #    field_next_if = builder.insert_field(aw.fields.FieldType.FIELD_NEXT_IF, True).as_field_next_if()
    #    field_next_if.left_expression = "5"
    #    field_next_if.right_expression = "2 + 3"
    #    field_next_if.comparison_operator = "="

    #    self.assertEqual(" NEXTIF  5 = \"2 + 3\"", field_next_if.get_field_code())

    #    # If the comparison asserted by the above field is correct,
    #    # the following 3 merge fields will take data from the third row.
    #    # Otherwise, these fields will take data from row 2 again.
    #    ExField.insert_merge_fields(builder, "Third row: ")

    #    doc.mail_merge.execute(table)

    #    # Our data source has 3 rows, and we skipped rows twice.
    #    # Our output document will have 1 page with data from all 3 rows.
    #    doc.save(ARTIFACTS_DIR + "Field.next.nextif.docx")
    #    self._test_field_next(doc) #ExSKip

    #@staticmethod
    #def insert_merge_fields(builder: aw.DocumentBuilder, first_field_text_before: str):
    #    """Uses a document builder to insert MERGEFIELDs for a data source that contains columns named "Courtesy Title", "First Name" and "Last Name"."""

    #    ExField.insert_merge_field(builder, "Courtesy Title", first_field_text_before, " ")
    #    ExField.insert_merge_field(builder, "First Name", None, " ")
    #    ExField.insert_merge_field(builder, "Last Name", None, None)
    #    builder.insert_paragraph()

    #@staticmethod
    #def insert_merge_field(builder: aw.DocumentBuilder, field_name: str, text_before: str, text_after: str):
    #    """Uses a document builder to insert a MERRGEFIELD with specified properties."""

    #    field = builder.insert_field(aw.fields.FieldType.FIELD_MERGE_FIELD, True).as_field_merge_field()
    #    field.field_name = field_name
    #    field.text_before = text_before
    #    field.text_after = text_after

    ##ExEnd

    def _test_field_next(self, doc: aw.Document):

        doc = DocumentHelper.save_open(doc)

        self.assertEqual(0, doc.range.fields.count)
        self.assertEqual("First row: Mr. John Doe\r" +
                        "Second row: Mrs. Jane Cardholder\r" +
                        "Third row: Mr. Joe Bloggs\r\f", doc.get_text())

    #ExStart
    #ExFor:FieldNoteRef
    #ExFor:FieldNoteRef.bookmark_name
    #ExFor:FieldNoteRef.insert_hyperlink
    #ExFor:FieldNoteRef.insert_reference_mark
    #ExFor:FieldNoteRef.insert_relative_position
    #ExSummary:Shows to insert NOTEREF fields, and modify their appearance.
    @unittest.skip("WORDSNET-17845") #ExSkip
    def test_field_note_ref(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create a bookmark with a footnote that the NOTEREF field will reference.
        ExField.insert_bookmark_with_footnote(builder, "MyBookmark1", "Contents of MyBookmark1", "Footnote from MyBookmark1")

        # This NOTEREF field will display the number of the footnote inside the referenced bookmark.
        # Setting the "insert_hyperlink" property lets us jump to the bookmark by Ctrl + clicking the field in Microsoft Word.
        self.assertEqual(" NOTEREF  MyBookmark2 \\h",
            ExField.insert_field_note_ref(builder, "MyBookmark2", True, False, False, "Hyperlink to Bookmark2, with footnote number ").get_field_code())

        # When using the \p flag, after the footnote number, the field also displays the bookmark's position relative to the field.
        # Bookmark1 is above this field and contains footnote number 1, so the result will be "1 above" on update.
        self.assertEqual(" NOTEREF  MyBookmark1 \\h \\p",
            ExField.insert_field_note_ref(builder, "MyBookmark1", True, True, False, "Bookmark1, with footnote number ").get_field_code())

        # Bookmark2 is below this field and contains footnote number 2, so the field will display "2 below".
        # The \f flag makes the number 2 appear in the same format as the footnote number label in the actual text.
        self.assertEqual(" NOTEREF  MyBookmark2 \\h \\p \\f",
            ExField.insert_field_note_ref(builder, "MyBookmark2", True, True, True, "Bookmark2, with footnote number ").get_field_code())

        builder.insert_break(aw.BreakType.PAGE_BREAK)
        ExField.insert_bookmark_with_footnote(builder, "MyBookmark2", "Contents of MyBookmark2", "Footnote from MyBookmark2")

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_note_ref.docx")
        self._test_note_ref(aw.Document(ARTIFACTS_DIR + "Field.field_note_ref.docx")) #ExSkip

    @staticmethod
    def insert_field_note_ref(builder: aw.DocumentBuilder, bookmark_name: str, insert_hyperlink: bool, insert_relative_position: bool, insert_reference_mark: bool, text_before: str) -> aw.fields.FieldNoteRef:
        """Uses a document builder to insert a NOTEREF field with specified properties."""

        builder.write(text_before)

        field = builder.insert_field(aw.fields.FieldType.FIELD_NOTE_REF, True).as_field_note_ref()
        field.bookmark_name = bookmark_name
        field.insert_hyperlink = insert_hyperlink
        field.insert_relative_position = insert_relative_position
        field.insert_reference_mark = insert_reference_mark
        builder.writeln()

        return field

    @staticmethod
    def insert_bookmark_with_footnote(builder: aw.DocumentBuilder, bookmark_name: str, bookmark_text: str, footnote_text: str):
        """Uses a document builder to insert a named bookmark with a footnote at the end."""

        builder.start_bookmark(bookmark_name)
        builder.write(bookmark_text)
        builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, footnote_text)
        builder.end_bookmark(bookmark_name)
        builder.writeln()

    #ExEnd

    def _test_note_ref(self, doc: aw.Document):

        field = doc.range.fields[0].as_field_note_ref()

        self.verify_field(aw.fields.FieldType.FIELD_NOTE_REF, " NOTEREF  MyBookmark2 \\h", "2", field)
        self.assertEqual("MyBookmark2", field.bookmark_name)
        self.assertTrue(field.insert_hyperlink)
        self.assertFalse(field.insert_relative_position)
        self.assertFalse(field.insert_reference_mark)

        field = doc.range.fields[1].as_field_note_ref()

        self.verify_field(aw.fields.FieldType.FIELD_NOTE_REF, " NOTEREF  MyBookmark1 \\h \\p", "1 above", field)
        self.assertEqual("MyBookmark1", field.bookmark_name)
        self.assertTrue(field.insert_hyperlink)
        self.assertTrue(field.insert_relative_position)
        self.assertFalse(field.insert_reference_mark)

        field = doc.range.fields[2].as_field_note_ref()

        self.verify_field(aw.fields.FieldType.FIELD_NOTE_REF, " NOTEREF  MyBookmark2 \\h \\p \\f", "2 below", field)
        self.assertEqual("MyBookmark2", field.bookmark_name)
        self.assertTrue(field.insert_hyperlink)
        self.assertTrue(field.insert_relative_position)
        self.assertTrue(field.insert_reference_mark)

    @unittest.skip("WORDSNET-17845")
    def test_field_footnote_ref(self):

        #ExStart
        #ExFor:FieldFootnoteRef
        #ExSummary:Shows how to cross-reference footnotes with the FOOTNOTEREF field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_bookmark("CrossRefBookmark")
        builder.write("Hello world!")
        builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "Cross referenced footnote.")
        builder.end_bookmark("CrossRefBookmark")
        builder.insert_paragraph()

        # Insert a FOOTNOTEREF field, which lets us reference a footnote more than once while re-using the same footnote marker.
        builder.write("CrossReference: ")
        field = builder.insert_field(aw.fields.FieldType.FIELD_FOOTNOTE_REF, True).as_field_footnote_ref()

        # Reference the bookmark that we have created with the FOOTNOTEREF field. That bookmark contains a footnote marker
        # belonging to the footnote we inserted. The field will display that footnote marker.
        builder.move_to(field.separator)
        builder.write("CrossRefBookmark")

        self.assertEqual(" FOOTNOTEREF CrossRefBookmark", field.get_field_code())

        doc.update_fields()

        # This field works only in older versions of Microsoft Word.
        doc.save(ARTIFACTS_DIR + "Field.field_footnote_ref.doc")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_footnote_ref.doc")
        field = doc.range.fields[0].as_field_footnote_ref()

        self.verify_field(aw.fields.FieldType.FIELD_FOOTNOTE_REF, " FOOTNOTEREF CrossRefBookmark", "1", field)
        self.verify_footnote(aw.notes.FootnoteType.FOOTNOTE, True, "", "Cross referenced footnote.",
            doc.get_child(aw.NodeType.FOOTNOTE, 0, True).as_footnote())

    #ExStart
    #ExFor:FieldPageRef
    #ExFor:FieldPageRef.bookmark_name
    #ExFor:FieldPageRef.insert_hyperlink
    #ExFor:FieldPageRef.insert_relative_position
    #ExSummary:Shows to insert PAGEREF fields to display the relative location of bookmarks.
    @unittest.skip("WORDSNET-17836") #ExSkip
    def test_field_page_ref(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        ExField.insert_and_name_bookmark(builder, "MyBookmark1")

        # Insert a PAGEREF field that displays what page a bookmark is on.
        # Set the InsertHyperlink flag to make the field also function as a clickable link to the bookmark.
        self.assertEqual(" PAGEREF  MyBookmark3 \\h",
            ExField.insert_field_page_ref(builder, "MyBookmark3", True, False, "Hyperlink to Bookmark3, on page: ").get_field_code())

        # We can use the \p flag to get the PAGEREF field to display
        # the bookmark's position relative to the position of the field.
        # Bookmark1 is on the same page and above this field, so this field's displayed result will be "above".
        self.assertEqual(" PAGEREF  MyBookmark1 \\h \\p",
            ExField.insert_field_page_ref(builder, "MyBookmark1", True, True, "Bookmark1 is ").get_field_code())

        # Bookmark2 will be on the same page and below this field, so this field's displayed result will be "below".
        self.assertEqual(" PAGEREF  MyBookmark2 \\h \\p",
            ExField.insert_field_page_ref(builder, "MyBookmark2", True, True, "Bookmark2 is ").get_field_code())

        # Bookmark3 will be on a different page, so the field will display "on page 2".
        self.assertEqual(" PAGEREF  MyBookmark3 \\h \\p",
            ExField.insert_field_page_ref(builder, "MyBookmark3", True, True, "Bookmark3 is ").get_field_code())

        ExField.insert_and_name_bookmark(builder, "MyBookmark2")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        ExField.insert_and_name_bookmark(builder, "MyBookmark3")

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_page_ref.docx")
        self._test_page_ref(aw.Document(ARTIFACTS_DIR + "Field.field_page_ref.docx")) #ExSkip

    @staticmethod
    def insert_field_page_ref(builder: aw.DocumentBuilder, bookmark_name: str, insert_hyperlink: bool, insert_relative_position: bool, text_before: str) -> aw.fields.FieldPageRef:
        """Uses a document builder to insert a PAGEREF field and sets its properties."""

        builder.write(text_before)

        field = builder.insert_field(aw.fields.FieldType.FIELD_PAGE_REF, True).as_field_page_ref()
        field.bookmark_name = bookmark_name
        field.insert_hyperlink = insert_hyperlink
        field.insert_relative_position = insert_relative_position
        builder.writeln()

        return field

    @staticmethod
    def insert_and_name_bookmark(builder: aw.DocumentBuilder, bookmark_name: str):
        """Uses a document builder to insert a named bookmark."""

        builder.start_bookmark(bookmark_name)
        builder.writeln(f"Contents of bookmark \"{bookmark_name}\".")
        builder.end_bookmark(bookmark_name)

    #ExEnd

    def _test_page_ref(self, doc: aw.Document):

        field = doc.range.fields[0].as_field_page_ref()

        self.verify_field(aw.fields.FieldType.FIELD_PAGE_REF, " PAGEREF  MyBookmark3 \\h", "2", field)
        self.assertEqual("MyBookmark3", field.bookmark_name)
        self.assertTrue(field.insert_hyperlink)
        self.assertFalse(field.insert_relative_position)

        field = doc.range.fields[1].as_field_page_ref()

        self.verify_field(aw.fields.FieldType.FIELD_PAGE_REF, " PAGEREF  MyBookmark1 \\h \\p", "above", field)
        self.assertEqual("MyBookmark1", field.bookmark_name)
        self.assertTrue(field.insert_hyperlink)
        self.assertTrue(field.insert_relative_position)

        field = doc.range.fields[2].as_field_page_ref()

        self.verify_field(aw.fields.FieldType.FIELD_PAGE_REF, " PAGEREF  MyBookmark2 \\h \\p", "below", field)
        self.assertEqual("MyBookmark2", field.bookmark_name)
        self.assertTrue(field.insert_hyperlink)
        self.assertTrue(field.insert_relative_position)

        field = doc.range.fields[3].as_field_page_ref()

        self.verify_field(aw.fields.FieldType.FIELD_PAGE_REF, " PAGEREF  MyBookmark3 \\h \\p", "on page 2", field)
        self.assertEqual("MyBookmark3", field.bookmark_name)
        self.assertTrue(field.insert_hyperlink)
        self.assertTrue(field.insert_relative_position)

    #ExStart
    #ExFor:FieldRef
    #ExFor:FieldRef.bookmark_name
    #ExFor:FieldRef.include_note_or_comment
    #ExFor:FieldRef.insert_hyperlink
    #ExFor:FieldRef.insert_paragraph_number
    #ExFor:FieldRef.insert_paragraph_number_in_full_context
    #ExFor:FieldRef.insert_paragraph_number_in_relative_context
    #ExFor:FieldRef.insert_relative_position
    #ExFor:FieldRef.number_separator
    #ExFor:FieldRef.suppress_non_delimiters
    #ExSummary:Shows how to insert REF fields to reference bookmarks.
    @unittest.skip("WORDSNET-18067") #ExSkip
    def test_field_ref(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_bookmark("MyBookmark")
        builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "MyBookmark footnote #1")
        builder.write("Text that will appear in REF field")
        builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "MyBookmark footnote #2")
        builder.end_bookmark("MyBookmark")
        builder.move_to_document_start()

        # We will apply a custom list format, where the amount of angle brackets indicates the list level we are currently at.
        builder.list_format.apply_number_default()
        builder.list_format.list_level.number_format = "> \x0000"

        # Insert a REF field that will contain the text within our bookmark, act as a hyperlink, and clone the bookmark's footnotes.
        field = ExField.insert_field_ref(builder, "MyBookmark", "", "\n")
        field.include_note_or_comment = True
        field.insert_hyperlink = True

        self.assertEqual(" REF  MyBookmark \\f \\h", field.get_field_code())

        # Insert a REF field, and display whether the referenced bookmark is above or below it.
        field = ExField.insert_field_ref(builder, "MyBookmark", "The referenced paragraph is ", " this field.\n")
        field.insert_relative_position = True

        self.assertEqual(" REF  MyBookmark \\p", field.get_field_code())

        # Display the list number of the bookmark as it appears in the document.
        field = ExField.insert_field_ref(builder, "MyBookmark", "The bookmark's paragraph number is ", "\n")
        field.insert_paragraph_number = True

        self.assertEqual(" REF  MyBookmark \\n", field.get_field_code())

        # Display the bookmark's list number, but with non-delimiter characters, such as the angle brackets, omitted.
        field = ExField.insert_field_ref(builder, "MyBookmark", "The bookmark's paragraph number, non-delimiters suppressed, is ", "\n")
        field.insert_paragraph_number = True
        field.suppress_non_delimiters = True

        self.assertEqual(" REF  MyBookmark \\n \\t", field.get_field_code())

        # Move down one list level.
        builder.list_format.list_level_number += 1
        builder.list_format.list_level.number_format = ">> \x0001"

        # Display the list number of the bookmark and the numbers of all the list levels above it.
        field = ExField.insert_field_ref(builder, "MyBookmark", "The bookmark's full context paragraph number is ", "\n")
        field.insert_paragraph_number_in_full_context = True

        self.assertEqual(" REF  MyBookmark \\w", field.get_field_code())

        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # Display the list level numbers between this REF field, and the bookmark that it is referencing.
        field = ExField.insert_field_ref(builder, "MyBookmark", "The bookmark's relative paragraph number is ", "\n")
        field.insert_paragraph_number_in_relative_context = True

        self.assertEqual(" REF  MyBookmark \\r", field.get_field_code())

        # At the end of the document, the bookmark will show up as a list item here.
        builder.writeln("List level above bookmark")
        builder.list_format.list_level_number += 1
        builder.list_format.list_level.number_format = ">>> \x0002"

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_ref.docx")
        self._test_field_ref(aw.Document(ARTIFACTS_DIR + "Field.field_ref.docx")) #ExSkip

    @staticmethod
    def insert_field_ref(builder: aw.DocumentBuilder, bookmark_name: str, text_before: str, text_after: str) -> aw.fields.FieldRef:
        """Get the document builder to insert a REF field, reference a bookmark with it, and add text before and after it."""

        builder.write(text_before)
        field = builder.insert_field(aw.fields.FieldType.FIELD_REF, True).as_field_ref()
        field.bookmark_name = bookmark_name
        builder.write(text_after)
        return field

    #ExEnd

    def _test_field_ref(self, doc: aw.Document):

        self.verify_footnote(aw.notes.FootnoteType.FOOTNOTE, True, "", "MyBookmark footnote #1",
            doc.get_child(aw.NodeType.FOOTNOTE, 0, True).as_footnote())
        self.verify_footnote(aw.notes.FootnoteType.FOOTNOTE, True, "", "MyBookmark footnote #2",
            doc.get_child(aw.NodeType.FOOTNOTE, 0, True).as_footnote())

        field = doc.range.fields[0].as_field_ref()

        self.verify_field(aw.fields.FieldType.FIELD_REF, " REF  MyBookmark \\f \\h",
            "\u0002 MyBookmark footnote #1\r" +
            "Text that will appear in REF field\u0002 MyBookmark footnote #2\r", field)
        self.assertEqual("MyBookmark", field.bookmark_name)
        self.assertTrue(field.include_note_or_comment)
        self.assertTrue(field.insert_hyperlink)

        field = doc.range.fields[1].as_field_ref()

        self.verify_field(aw.fields.FieldType.FIELD_REF, " REF  MyBookmark \\p", "below", field)
        self.assertEqual("MyBookmark", field.bookmark_name)
        self.assertTrue(field.insert_relative_position)

        field = doc.range.fields[2].as_field_ref()

        self.verify_field(aw.fields.FieldType.FIELD_REF, " REF  MyBookmark \\n", ">>> i", field)
        self.assertEqual("MyBookmark", field.bookmark_name)
        self.assertTrue(field.insert_paragraph_number)
        self.assertEqual(" REF  MyBookmark \\n", field.get_field_code())
        self.assertEqual(">>> i", field.result)

        field = doc.range.fields[3].as_field_ref()

        self.verify_field(aw.fields.FieldType.FIELD_REF, " REF  MyBookmark \\n \\t", "i", field)
        self.assertEqual("MyBookmark", field.bookmark_name)
        self.assertTrue(field.insert_paragraph_number)
        self.assertTrue(field.suppress_non_delimiters)

        field = doc.range.fields[4].as_field_ref()

        self.verify_field(aw.fields.FieldType.FIELD_REF, " REF  MyBookmark \\w", "> 4>> c>>> i", field)
        self.assertEqual("MyBookmark", field.bookmark_name)
        self.assertTrue(field.insert_paragraph_number_in_full_context)

        field = doc.range.fields[5].as_field_ref()

        self.verify_field(aw.fields.FieldType.FIELD_REF, " REF  MyBookmark \\r", ">> c>>> i", field)
        self.assertEqual("MyBookmark", field.bookmark_name)
        self.assertTrue(field.insert_paragraph_number_in_relative_context)

    @unittest.skip("WORDSNET-18068")
    def test_field_rd(self):

        #ExStart
        #ExFor:FieldRD
        #ExFor:FieldRD.file_name
        #ExFor:FieldRD.is_path_relative
        #ExSummary:Shows to use the RD field to create a table of contents entries from headings in other documents.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Use a document builder to insert a table of contents,
        # and then add one entry for the table of contents on the following page.
        builder.insert_field(aw.fields.FieldType.FIELD_TOC, True)
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.current_paragraph.paragraph_format.style_name = "Heading 1"
        builder.writeln("TOC entry from within this document")

        # Insert an RD field, which references another local file system document in its FileName property.
        # The TOC will also now accept all headings from the referenced document as entries for its table.
        field = builder.insert_field(aw.fields.FieldType.FIELD_REF_DOC, True).as_field_rd()
        field.file_name = "ReferencedDocument.docx"
        field.is_path_relative = True

        self.assertEqual(" RD  ReferencedDocument.docx \\f", field.get_field_code())

        # Create the document that the RD field is referencing and insert a heading.
        # This heading will show up as an entry in the TOC field in our first document.
        referenced_doc = aw.Document()
        ref_doc_builder = aw.DocumentBuilder(referenced_doc)
        ref_doc_builder.current_paragraph.paragraph_format.style_name = "Heading 1"
        ref_doc_builder.writeln("TOC entry from referenced document")
        referenced_doc.save(ARTIFACTS_DIR + "ReferencedDocument.docx")

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_rd.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_rd.docx")

        field_toc = doc.range.fields[0].as_field_toc()

        self.assertEqual("TOC entry from within this document\t\u0013 PAGEREF _Toc36149519 \\h \u00142\u0015\r" +
                        "TOC entry from referenced document\t1\r", field_toc.result)

        field_page_ref = doc.range.fields[1].as_field_page_ref()

        self.verify_field(aw.fields.FieldType.FIELD_PAGE_REF, " PAGEREF _Toc36149519 \\h ", "2", field_page_ref)

        field = doc.range.fields[2].as_field_rd()

        self.verify_field(aw.fields.FieldType.FIELD_REF_DOC, " RD  ReferencedDocument.docx \\f", "", field)
        self.assertEqual("ReferencedDocument.docx", field.file_name)
        self.assertTrue(field.is_path_relative)

    #def test_skip_if(self):

    #    #ExStart
    #    #ExFor:FieldSkipIf
    #    #ExFor:FieldSkipIf.comparison_operator
    #    #ExFor:FieldSkipIf.left_expression
    #    #ExFor:FieldSkipIf.right_expression
    #    #ExSummary:Shows how to skip pages in a mail merge using the SKIPIF field.
    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    # Insert a SKIPIF field. If the current row of a mail merge operation fulfills the condition
    #    # which the expressions of this field state, then the mail merge operation aborts the current row,
    #    # discards the current merge document, and then immediately moves to the next row to begin the next merge document.
    #    field_skip_if = builder.insert_field(aw.fields.FieldType.FIELD_SKIP_IF, True).as_field_skip_if()

    #    # Move the builder to the SKIPIF field's separator so we can place a MERGEFIELD inside the SKIPIF field.
    #    builder.move_to(field_skip_if.separator)
    #    field_merge_field = builder.insert_field(aw.fields.FieldType.FIELD_MERGE_FIELD, True).as_field_merge_field()
    #    field_merge_field.field_name = "Department"

    #    # The MERGEFIELD refers to the "Department" column in our data table. If a row from that table
    #    # has a value of "HR" in its "Department" column, then this row will fulfill the condition.
    #    field_skip_if.left_expression = "="
    #    field_skip_if.right_expression = "HR"

    #    # Add content to our document, create the data source, and execute the mail merge.
    #    builder.move_to_document_end()
    #    builder.write("Dear ")
    #    field_merge_field = builder.insert_field(aw.fields.FieldType.FIELD_MERGE_FIELD, True).as_field_merge_field()
    #    field_merge_field.field_name = "Name"
    #    builder.writeln(", ")

    #    # This table has three rows, and one of them fulfills the condition of our SKIPIF field.
    #    # The mail merge will produce two pages.
    #    table = DataTable("Employees")
    #    table.columns.add("Name")
    #    table.columns.add("Department")
    #    table.rows.add("John Doe", "Sales")
    #    table.rows.add("Jane Doe", "Accounting")
    #    table.rows.add("John Cardholder", "HR")

    #    doc.mail_merge.execute(table)
    #    doc.save(ARTIFACTS_DIR + "Field.skipif.docx")
    #    #ExEnd

    #    doc = aw.Document(ARTIFACTS_DIR + "Field.skipif.docx")

    #    self.assertEqual(0, doc.range.fields.count)
    #    self.assertEqual("Dear John Doe, \r" +
    #                    "\fDear Jane Doe, \r\f", doc.get_text())

    def test_field_set_ref(self):

        #ExStart
        #ExFor:FieldRef
        #ExFor:FieldRef.bookmark_name
        #ExFor:FieldSet
        #ExFor:FieldSet.bookmark_name
        #ExFor:FieldSet.bookmark_text
        #ExSummary:Shows how to create bookmarked text with a SET field, and then display it in the document using a REF field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Name bookmarked text with a SET field.
        # This field refers to the "bookmark" not a bookmark structure that appears within the text, but a named variable.
        field_set = builder.insert_field(aw.fields.FieldType.FIELD_SET, False).as_field_set()
        field_set.bookmark_name = "MyBookmark"
        field_set.bookmark_text = "Hello world!"
        field_set.update()

        self.assertEqual(" SET  MyBookmark \"Hello world!\"", field_set.get_field_code())

        # Refer to the bookmark by name in a REF field and display its contents.
        field_ref = builder.insert_field(aw.fields.FieldType.FIELD_REF, True).as_field_ref()
        field_ref.bookmark_name = "MyBookmark"
        field_ref.update()

        self.assertEqual(" REF  MyBookmark", field_ref.get_field_code())
        self.assertEqual("Hello world!", field_ref.result)

        doc.save(ARTIFACTS_DIR + "Field.field_set_ref.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_set_ref.docx")

        self.assertEqual("Hello world!", doc.range.bookmarks[0].text)

        field_set = doc.range.fields[0].as_field_set()

        self.verify_field(aw.fields.FieldType.FIELD_SET, " SET  MyBookmark \"Hello world!\"", "Hello world!", field_set)
        self.assertEqual("MyBookmark", field_set.bookmark_name)
        self.assertEqual("Hello world!", field_set.bookmark_text)

        self.verify_field(aw.fields.FieldType.FIELD_REF, " REF  MyBookmark", "Hello world!", field_ref)
        self.assertEqual("Hello world!", field_ref.result)

    def test_field_template(self):

        #ExStart
        #ExFor:FieldTemplate
        #ExFor:FieldTemplate.include_full_path
        #ExFor:FieldOptions.template_name
        #ExSummary:Shows how to use a TEMPLATE field to display the local file system location of a document's template.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # We can set a template name using by the fields. This property is used when the "doc.attached_template" is empty.
        # If this property is empty the default template file name "Normal.dotm" is used.
        doc.field_options.template_name = ""

        field = builder.insert_field(aw.fields.FieldType.FIELD_TEMPLATE, False).as_field_template()
        self.assertEqual(" TEMPLATE ", field.get_field_code())

        builder.writeln()
        field = builder.insert_field(aw.fields.FieldType.FIELD_TEMPLATE, False).as_field_template()
        field.include_full_path = True

        self.assertEqual(" TEMPLATE  \\p", field.get_field_code())

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_template.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_template.docx")

        field = doc.range.fields[0].as_field_template()
        self.assertEqual(" TEMPLATE ", field.get_field_code())
        self.assertEqual("Normal.dotm", field.result)

        field = doc.range.fields[1].as_field_template()
        self.assertEqual(" TEMPLATE  \\p", field.get_field_code())
        self.assertEqual("Normal.dotm", field.result)

    def test_field_symbol(self):

        #ExStart
        #ExFor:FieldSymbol
        #ExFor:FieldSymbol.character_code
        #ExFor:FieldSymbol.dont_affects_line_spacing
        #ExFor:FieldSymbol.font_name
        #ExFor:FieldSymbol.font_size
        #ExFor:FieldSymbol.is_ansi
        #ExFor:FieldSymbol.is_shift_jis
        #ExFor:FieldSymbol.is_unicode
        #ExSummary:Shows how to use the SYMBOL field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Below are three ways to use a SYMBOL field to display a single character.
        # 1 -  Add a SYMBOL field which displays the © (Copyright) symbol, specified by an ANSI character code:
        field = builder.insert_field(aw.fields.FieldType.FIELD_SYMBOL, True).as_field_symbol()

        # The ANSI character code "U+00A9", or "169" in integer form, is reserved for the copyright symbol.
        field.character_code = "\u00a9"
        field.is_ansi = True

        self.assertEqual(" SYMBOL  169 \\a", field.get_field_code())

        builder.writeln(" Line 1")

        # 2 -  Add a SYMBOL field which displays the ∞ (Infinity) symbol, and modify its appearance:
        field = builder.insert_field(aw.fields.FieldType.FIELD_SYMBOL, True).as_field_symbol()

        # In Unicode, the infinity symbol occupies the "221E" code.
        field.character_code = "/u221E"
        field.is_unicode = True

        # Change the font of our symbol after using the Windows Character Map
        # to ensure that the font can represent that symbol.
        field.font_name = "Calibri"
        field.font_size = "24"

        # We can set this flag for tall symbols to make them not push down the rest of the text on their line.
        field.dont_affects_line_spacing = True

        self.assertEqual(" SYMBOL  8734 \\u \\f Calibri \\s 24 \\h", field.get_field_code())

        builder.writeln("Line 2")

        # 3 -  Add a SYMBOL field which displays the あ character,
        # with a font that supports Shift-JIS (Windows-932) codepage:
        field = builder.insert_field(aw.fields.FieldType.FIELD_SYMBOL, True).as_field_symbol()
        field.font_name = "MS Gothic"
        field.character_code = "/u82A0"
        field.is_shift_jis = True

        self.assertEqual(" SYMBOL  33440 \\f \"MS Gothic\" \\j", field.get_field_code())

        builder.write("Line 3")

        doc.save(ARTIFACTS_DIR + "Field.field_symbol.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_symbol.docx")

        field = doc.range.fields[0].as_field_symbol()

        self.verify_field(aw.fields.FieldType.FIELD_SYMBOL, " SYMBOL  169 \\a", "", field)
        self.assertEqual(0x00a9.to_string(), field.character_code)
        self.assertTrue(field.is_ansi)
        self.assertEqual("©", field.display_result)

        field = doc.range.fields[1].as_field_symbol()

        self.verify_field(aw.fields.FieldType.FIELD_SYMBOL, " SYMBOL  8734 \\u \\f Calibri \\s 24 \\h", "", field)
        self.assertEqual(0x221E.to_string(), field.character_code)
        self.assertEqual("Calibri", field.font_name)
        self.assertEqual("24", field.font_size)
        self.assertTrue(field.is_unicode)
        self.assertTrue(field.dont_affects_line_spacing)
        self.assertEqual("∞", field.display_result)

        field = doc.range.fields[2].as_field_symbol()

        self.verify_field(aw.fields.FieldType.FIELD_SYMBOL, " SYMBOL  33440 \\f \"MS Gothic\" \\j", "", field)
        self.assertEqual(0x82A0.to_string(), field.character_code)
        self.assertEqual("MS Gothic", field.font_name)
        self.assertTrue(field.is_shift_jis)

    def test_field_title(self):

        #ExStart
        #ExFor:FieldTitle
        #ExFor:FieldTitle.text
        #ExSummary:Shows how to use the TITLE field.
        doc = aw.Document()

        # Set a value for the "title" built-in document property.
        doc.built_in_document_properties.title = "My Title"

        # We can use the TITLE field to display the value of this property in the document.
        builder = aw.DocumentBuilder(doc)
        field = builder.insert_field(aw.fields.FieldType.FIELD_TITLE, False).as_field_title()
        field.update()

        self.assertEqual(" TITLE ", field.get_field_code())
        self.assertEqual("My Title", field.result)

        # Setting a value for the field's "text" property,
        # and then updating the field will also overwrite the corresponding built-in property with the new value.
        builder.writeln()
        field = builder.insert_field(aw.fields.FieldType.FIELD_TITLE, False).as_field_title()
        field.text = "My New Title"
        field.update()

        self.assertEqual(" TITLE  \"My New Title\"", field.get_field_code())
        self.assertEqual("My New Title", field.result)
        self.assertEqual("My New Title", doc.built_in_document_properties.title)

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_title.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_title.docx")

        self.assertEqual("My New Title", doc.built_in_document_properties.title)

        field = doc.range.fields[0].as_field_title()

        self.verify_field(aw.fields.FieldType.FIELD_TITLE, " TITLE ", "My New Title", field)

        field = doc.range.fields[1].as_field_title()

        self.verify_field(aw.fields.FieldType.FIELD_TITLE, " TITLE  \"My New Title\"", "My New Title", field)
        self.assertEqual("My New Title", field.text)

    #ExStart
    #ExFor:FieldToa
    #ExFor:FieldToa.bookmark_name
    #ExFor:FieldToa.entry_category
    #ExFor:FieldToa.entry_separator
    #ExFor:FieldToa.page_number_list_separator
    #ExFor:FieldToa.page_range_separator
    #ExFor:FieldToa.remove_entry_formatting
    #ExFor:FieldToa.sequence_name
    #ExFor:FieldToa.sequence_separator
    #ExFor:FieldToa.use_heading
    #ExFor:FieldToa.use_passim
    #ExFor:FieldTA
    #ExFor:FieldTA.entry_category
    #ExFor:FieldTA.is_bold
    #ExFor:FieldTA.is_italic
    #ExFor:FieldTA.long_citation
    #ExFor:FieldTA.page_range_bookmark_name
    #ExFor:FieldTA.short_citation
    #ExSummary:Shows how to build and customize a table of authorities using TOA and TA fields.
    def test_field_toa(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a TOA field, which will create an entry for each TA field in the document,
        # displaying long citations and page numbers for each entry.
        field_toa = builder.insert_field(aw.fields.FieldType.FIELD_TOA, False).as_field_toa()

        # Set the entry category for our table. This TOA will now only include TA fields
        # that have a matching value in their "entry_category" property.
        field_toa.entry_category = "1"

        # Moreover, the Table of Authorities category at index 1 is "Cases",
        # which will show up as our table's title if we set this variable to True.
        field_toa.use_heading = True

        # We can further filter TA fields by naming a bookmark that they will need to be within the TOA bounds.
        field_toa.bookmark_name = "MyBookmark"

        # By default, a dotted line page-wide tab appears between the TA field's citation
        # and its page number. We can replace it with any text we put on this property.
        # Inserting a tab character will preserve the original tab.
        field_toa.entry_separator = " \t p."

        # If we have multiple TA entries that share the same long citation,
        # all their respective page numbers will show up on one row.
        # We can use this property to specify a string that will separate their page numbers.
        field_toa.page_number_list_separator = " & p. "

        # We can set this to True to get our table to display the word "passim"
        # if there are five or more page numbers in one row.
        field_toa.use_passim = True

        # One TA field can refer to a range of pages.
        # We can specify a string here to appear between the start and end page numbers for such ranges.
        field_toa.page_range_separator = " to "

        # The format from the TA fields will carry over into our table.
        # We can disable this by setting the "remove_entry_formatting" flag.
        field_toa.remove_entry_formatting = True
        builder.font.color = drawing.Color.green
        builder.font.name = "Arial Black"

        self.assertEqual(" TOA  \\c 1 \\h \\b MyBookmark \\e \" \t p.\" \\l \" & p. \" \\p \\g \" to \" \\f", field_toa.get_field_code())

        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # This TA field will not appear as an entry in the TOA since it is outside
        # the bookmark's bounds that the TOA's "bookmark_name" property specifies.
        field_ta = ExField.insert_toa_entry(builder, "1", "Source 1")

        self.assertEqual(" TA  \\c 1 \\l \"Source 1\"", field_ta.get_field_code())

        # This TA field is inside the bookmark,
        # but the entry category does not match that of the table, so the TA field will not include it.
        builder.start_bookmark("MyBookmark")
        field_ta = ExField.insert_toa_entry(builder, "2", "Source 2")

        # This entry will appear in the table.
        field_ta = ExField.insert_toa_entry(builder, "1", "Source 3")

        # A TOA table does not display short citations,
        # but we can use them as a shorthand to refer to bulky source names that multiple TA fields reference.
        field_ta.short_citation = "S.3"

        self.assertEqual(" TA  \\c 1 \\l \"Source 3\" \\s S.3", field_ta.get_field_code())

        # We can format the page number to make it bold/italic using the following properties.
        # We will still see these effects if we set our table to ignore formatting.
        field_ta = ExField.insert_toa_entry(builder, "1", "Source 2")
        field_ta.is_bold = True
        field_ta.is_italic = True

        self.assertEqual(" TA  \\c 1 \\l \"Source 2\" \\b \\i", field_ta.get_field_code())

        # We can configure TA fields to get their TOA entries to refer to a range of pages that a bookmark spans across.
        # Note that this entry refers to the same source as the one above to share one row in our table.
        # This row will have the page number of the entry above and the page range of this entry,
        # with the table's page list and page number range separators between page numbers.
        field_ta = ExField.insert_toa_entry(builder, "1", "Source 3")
        field_ta.page_range_bookmark_name = "MyMultiPageBookmark"

        builder.start_bookmark("MyMultiPageBookmark")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.end_bookmark("MyMultiPageBookmark")

        self.assertEqual(" TA  \\c 1 \\l \"Source 3\" \\r MyMultiPageBookmark", field_ta.get_field_code())

        # If we have enabled the "Passim" feature of our table, having 5 or more TA entries with the same source will invoke it.
        for i in range(5):
            ExField.insert_toa_entry(builder, "1", "Source 4")

        builder.end_bookmark("MyBookmark")

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_toa.docx")
        self._test_field_toa(aw.Document(ARTIFACTS_DIR + "Field.field_toa.docx")) #ExSKip

    @staticmethod
    def insert_toa_entry(builder: aw.DocumentBuilder, entry_category: str, long_citation: str) -> aw.fields.FieldTA:

        field = builder.insert_field(aw.fields.FieldType.FIELD_TOAENTRY, False).as_field_ta()
        field.entry_category = entry_category
        field.long_citation = long_citation

        builder.insert_break(aw.BreakType.PAGE_BREAK)

        return field

    #ExEnd

    def _test_field_toa(self, doc: aw.Document):

        field_toa = doc.range.fields[0].as_field_toa()

        self.assertEqual("1", field_toa.entry_category)
        self.assertTrue(field_toa.use_heading)
        self.assertEqual("MyBookmark", field_toa.bookmark_name)
        self.assertEqual(" \t p.", field_toa.entry_separator)
        self.assertEqual(" & p. ", field_toa.page_number_list_separator)
        self.assertTrue(field_toa.use_passim)
        self.assertEqual(" to ", field_toa.page_range_separator)
        self.assertTrue(field_toa.remove_entry_formatting)
        self.assertEqual(" TOA  \\c 1 \\h \\b MyBookmark \\e \" \t p.\" \\l \" & p. \" \\p \\g \" to \" \\f", field_toa.get_field_code())
        self.assertEqual("Cases\r" +
                        "Source 2 \t p.5\r" +
                        "Source 3 \t p.4 & p. 7 to 10\r" +
                        "Source 4 \t p.passim\r", field_toa.result)

        field_ta = doc.range.fields[1].as_field_ta()

        self.verify_field(aw.fields.FieldType.FIELD_TOAENTRY, " TA  \\c 1 \\l \"Source 1\"", "", field_ta)
        self.assertEqual("1", field_ta.entry_category)
        self.assertEqual("Source 1", field_ta.long_citation)

        field_ta = doc.range.fields[2].as_field_ta()

        self.verify_field(aw.fields.FieldType.FIELD_TOAENTRY, " TA  \\c 2 \\l \"Source 2\"", "", field_ta)
        self.assertEqual("2", field_ta.entry_category)
        self.assertEqual("Source 2", field_ta.long_citation)

        field_ta = doc.range.fields[3].as_field_ta()

        self.verify_field(aw.fields.FieldType.FIELD_TOAENTRY, " TA  \\c 1 \\l \"Source 3\" \\s S.3", "", field_ta)
        self.assertEqual("1", field_ta.entry_category)
        self.assertEqual("Source 3", field_ta.long_citation)
        self.assertEqual("S.3", field_ta.short_citation)

        field_ta = doc.range.fields[4].as_field_ta()

        self.verify_field(aw.fields.FieldType.FIELD_TOAENTRY, " TA  \\c 1 \\l \"Source 2\" \\b \\i", "", field_ta)
        self.assertEqual("1", field_ta.entry_category)
        self.assertEqual("Source 2", field_ta.long_citation)
        self.assertTrue(field_ta.is_bold)
        self.assertTrue(field_ta.is_italic)

        field_ta = doc.range.fields[5].as_field_ta()

        self.verify_field(aw.fields.FieldType.FIELD_TOAENTRY, " TA  \\c 1 \\l \"Source 3\" \\r MyMultiPageBookmark", "", field_ta)
        self.assertEqual("1", field_ta.entry_category)
        self.assertEqual("Source 3", field_ta.long_citation)
        self.assertEqual("MyMultiPageBookmark", field_ta.page_range_bookmark_name)

        for i in range(6, 11):
            field_ta = doc.range.fields[i].as_field_ta()

            self.verify_field(aw.fields.FieldType.FIELD_TOAENTRY, " TA  \\c 1 \\l \"Source 4\"", "", field_ta)
            self.assertEqual("1", field_ta.entry_category)
            self.assertEqual("Source 4", field_ta.long_citation)

    def test_field_add_in(self):

        #ExStart
        #ExFor:FieldAddIn
        #ExSummary:Shows how to process an ADDIN field.
        doc = aw.Document(MY_DIR + "Field sample - ADDIN.docx")

        # Aspose.Words does not support inserting ADDIN fields, but we can still load and read them.
        field = doc.range.fields[0].as_field_add_in()

        self.assertEqual(" ADDIN \"My value\" ", field.get_field_code())
        #ExEnd

        doc = DocumentHelper.save_open(doc)

        self.verify_field(aw.fields.FieldType.FIELD_ADDIN, " ADDIN \"My value\" ", "", doc.range.fields[0])

    def test_field_edit_time(self):

        #ExStart
        #ExFor:FieldEditTime
        #ExSummary:Shows how to use the EDITTIME field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # The EDITTIME field will show, in minutes,
        # the time spent with the document open in a Microsoft Word window.
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.write("You've been editing this document for ")
        field = builder.insert_field(aw.fields.FieldType.FIELD_EDIT_TIME, True).as_field_edit_time()
        builder.writeln(" minutes.")

        # This built in document property tracks the minutes. Microsoft Word uses this property
        # to track the time spent with the document open. We can also edit it ourselves.
        doc.built_in_document_properties.total_editing_time = 10
        field.update()

        self.assertEqual(" EDITTIME ", field.get_field_code())
        self.assertEqual("10", field.result)

        # The field does not update itself in real-time, and will also have to be
        # manually updated in Microsoft Word anytime we need an accurate value.
        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_edit_time.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_edit_time.docx")

        self.assertEqual(10, doc.built_in_document_properties.total_editing_time)

        self.verify_field(aw.fields.FieldType.FIELD_EDIT_TIME, " EDITTIME ", "10", doc.range.fields[0])

    #ExStart
    #ExFor:FieldEQ
    #ExSummary:Shows how to use the EQ field to display a variety of mathematical equations.
    def test_field_eq(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # An EQ field displays a mathematical equation consisting of one or many elements.
        # Each element takes the following form: [switch][options][arguments].
        # There may be one switch, and several possible options.
        # The arguments are a set of coma-separated values enclosed by round braces.

        # Here we use a document builder to insert an EQ field, with an "\f" switch, which corresponds to "Fraction".
        # We will pass values 1 and 4 as arguments, and we will not use any options.
        # This field will display a fraction with 1 as the numerator and 4 as the denominator.
        field = ExField.insert_field_eq(builder, r"\f(1,4)")

        self.assertEqual(r" EQ \f(1,4)", field.get_field_code())

        # One EQ field may contain multiple elements placed sequentially.
        # We can also nest elements inside one another by placing the inner elements
        # inside the argument brackets of outer elements.
        # We can find the full list of switches, along with their uses here:
        # https://blogs.msdn.microsoft.com/murrays/2018/01/23/microsoft-word-eq-field/

        # Below are applications of nine different EQ field switches that we can use to create different kinds of objects.
        # 1 -  Array switch "\a", aligned left, 2 columns, 3 points of horizontal and vertical spacing:
        ExField.insert_field_eq(builder, r"\a \al \co2 \vs3 \hs3(4x,- 4y,-4x,+ y)")

        # 2 -  Bracket switch "\b", bracket character "[", to enclose the contents in a set of square braces:
        # Note that we are nesting an array inside the brackets, which will altogether look like a matrix in the output.
        ExField.insert_field_eq(builder, r"\b \bc\[ (\a \al \co3 \vs3 \hs3(1,0,0,0,1,0,0,0,1))")

        # 3 -  Displacement switch "\d", displacing text "B" 30 spaces to the right of "A", displaying the gap as an underline:
        ExField.insert_field_eq(builder, r"A \d \fo30 \li() B")

        # 4 -  Formula consisting of multiple fractions:
        ExField.insert_field_eq(builder, r"\f(d,dx)(u + v) = \f(du,dx) + \f(dv,dx)")

        # 5 -  Integral switch "\i", with a summation symbol:
        ExField.insert_field_eq(builder, r"\i \su(n=1,5,n)")

        # 6 -  List switch "\l":
        ExField.insert_field_eq(builder, r"\l(1,1,2,3,n,8,13)")

        # 7 -  Radical switch "\r", displaying a cubed root of x:
        ExField.insert_field_eq(builder, r"\r (3,x)")

        # 8 -  Subscript/superscript switch "/s", first as a superscript and then as a subscript:
        ExField.insert_field_eq(builder, r"\s \up8(Superscript) Text \s \do8(Subscript)")

        # 9 -  Box switch "\x", with lines at the top, bottom, left and right of the input:
        ExField.insert_field_eq(builder, r"\x \to \bo \le \ri(5)")

        # Some more complex combinations.
        ExField.insert_field_eq(builder, r"\a \ac \vs1 \co1(lim,n→∞) \b (\f(n,n2 + 12) + \f(n,n2 + 22) + ... + \f(n,n2 + n2))")
        ExField.insert_field_eq(builder, r"\i (,,  \b(\f(x,x2 + 3x + 2))) \s \up10(2)")
        ExField.insert_field_eq(builder, r"\i \in( tan x, \s \up2(sec x), \b(\r(3) )\s \up4(t) \s \up7(2)  dt)")

        doc.save(ARTIFACTS_DIR + "Field.field_eq.docx")
        self._test_field_eq(aw.Document(ARTIFACTS_DIR + "Field.field_eq.docx")) #ExSkip

    @staticmethod
    def insert_field_eq(builder: aw.DocumentBuilder, args: str) -> aw.fields.FieldEQ:
        """Use a document builder to insert an EQ field, set its arguments and start a new paragraph."""

        field = builder.insert_field(aw.fields.FieldType.FIELD_EQUATION, True).as_field_eq()
        builder.move_to(field.separator)
        builder.write(args)
        builder.move_to(field.start.parent_node)

        builder.insert_paragraph()
        return field

    #ExEnd

    def _test_field_eq(self, doc: aw.Document):

        self.verify_field(aw.fields.FieldType.FIELD_EQUATION, r" EQ \f(1,4)", "", doc.range.fields[0])
        self.verify_field(aw.fields.FieldType.FIELD_EQUATION, r" EQ \a \al \co2 \vs3 \hs3(4x,- 4y,-4x,+ y)", "", doc.range.fields[1])
        self.verify_field(aw.fields.FieldType.FIELD_EQUATION, r" EQ \b \bc\[ (\a \al \co3 \vs3 \hs3(1,0,0,0,1,0,0,0,1))", "", doc.range.fields[2])
        self.verify_field(aw.fields.FieldType.FIELD_EQUATION, r" EQ A \d \fo30 \li() B", "", doc.range.fields[3])
        self.verify_field(aw.fields.FieldType.FIELD_EQUATION, r" EQ \f(d,dx)(u + v) = \f(du,dx) + \f(dv,dx)", "", doc.range.fields[4])
        self.verify_field(aw.fields.FieldType.FIELD_EQUATION, r" EQ \i \su(n=1,5,n)", "", doc.range.fields[5])
        self.verify_field(aw.fields.FieldType.FIELD_EQUATION, r" EQ \l(1,1,2,3,n,8,13)", "", doc.range.fields[6])
        self.verify_field(aw.fields.FieldType.FIELD_EQUATION, r" EQ \r (3,x)", "", doc.range.fields[7])
        self.verify_field(aw.fields.FieldType.FIELD_EQUATION, r" EQ \s \up8(Superscript) Text \s \do8(Subscript)", "", doc.range.fields[8])
        self.verify_field(aw.fields.FieldType.FIELD_EQUATION, r" EQ \x \to \bo \le \ri(5)", "", doc.range.fields[9])
        self.verify_field(aw.fields.FieldType.FIELD_EQUATION, r" EQ \a \ac \vs1 \co1(lim,n→∞) \b (\f(n,n2 + 12) + \f(n,n2 + 22) + ... + \f(n,n2 + n2))", "", doc.range.fields[10])
        self.verify_field(aw.fields.FieldType.FIELD_EQUATION, r" EQ \i (,,  \b(\f(x,x2 + 3x + 2))) \s \up10(2)", "", doc.range.fields[11])
        self.verify_field(aw.fields.FieldType.FIELD_EQUATION, r" EQ \i \in( tan x, \s \up2(sec x), \b(\r(3) )\s \up4(t) \s \up7(2)  dt)", "", doc.range.fields[12])
        self.verify_web_response_status_code(200, "https://blogs.msdn.microsoft.com/murrays/2018/01/23/microsoft-word-eq-field/")

    def test_field_forms(self):

        #ExStart
        #ExFor:FieldFormCheckBox
        #ExFor:FieldFormDropDown
        #ExFor:FieldFormText
        #ExSummary:Shows how to process FORMCHECKBOX, FORMDROPDOWN and FORMTEXT fields.
        # These fields are legacy equivalents of the FormField. We can read, but not create these fields using Aspose.Words.
        # In Microsoft Word, we can insert these fields via the Legacy Tools menu in the Developer tab.
        doc = aw.Document(MY_DIR + "Form fields.docx")

        field_form_check_box = doc.range.fields[1].as_field_form_check_box()
        self.assertEqual(" FORMCHECKBOX \u0001", field_form_check_box.get_field_code())

        field_form_drop_down = doc.range.fields[2].as_field_form_drop_down()
        self.assertEqual(" FORMDROPDOWN \u0001", field_form_drop_down.get_field_code())

        field_form_text = doc.range.fields[0].as_field_form_text()
        self.assertEqual(" FORMTEXT \u0001", field_form_text.get_field_code())
        #ExEnd

    def test_field_formula(self):

        #ExStart
        #ExFor:FieldFormula
        #ExSummary:Shows how to use the formula field to display the result of an equation.
        doc = aw.Document()

        # Use a field builder to construct a mathematical equation,
        # then create a formula field to display the equation's result in the document.
        field_builder = aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_FORMULA)
        field_builder.add_argument(2)
        field_builder.add_argument("*")
        field_builder.add_argument(5)

        field = field_builder.build_and_insert(doc.first_section.body.first_paragraph).as_field_formula()
        field.update()

        self.assertEqual(" = 2 * 5 ", field.get_field_code())
        self.assertEqual("10", field.result)

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_formula.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_formula.docx")

        self.verify_field(aw.fields.FieldType.FIELD_FORMULA, " = 2 * 5 ", "10", doc.range.fields[0])

    def test_field_last_saved_by(self):

        #ExStart
        #ExFor:FieldLastSavedBy
        #ExSummary:Shows how to use the LASTSAVEDBY field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # If we create a document in Microsoft Word, it will have the user's name in the "Last saved by" built-in property.
        # If we make a document programmatically, this property will be null, and we will need to assign a value.
        doc.built_in_document_properties.last_saved_by = "John Doe"

        # We can use the LASTSAVEDBY field to display the value of this property in the document.
        field = builder.insert_field(aw.fields.FieldType.FIELD_LAST_SAVED_BY, True).as_field_last_saved_by()

        self.assertEqual(" LASTSAVEDBY ", field.get_field_code())
        self.assertEqual("John Doe", field.result)

        doc.save(ARTIFACTS_DIR + "Field.field_last_saved_by.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_last_saved_by.docx")

        self.assertEqual("John Doe", doc.built_in_document_properties.last_saved_by)
        self.verify_field(aw.fields.FieldType.FIELD_LAST_SAVED_BY, " LASTSAVEDBY ", "John Doe", doc.range.fields[0])

    #@unittest.skip("WORDSNET-18173")
    #def test_field_merge_rec(self):

    #    #ExStart
    #    #ExFor:FieldMergeRec
    #    #ExFor:FieldMergeSeq
    #    #ExFor:FieldSkipIf
    #    #ExFor:FieldSkipIf.comparison_operator
    #    #ExFor:FieldSkipIf.left_expression
    #    #ExFor:FieldSkipIf.right_expression
    #    #ExSummary:Shows how to use MERGEREC and MERGESEQ fields to the number and count mail merge records in a mail merge's output documents.
    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    builder.write("Dear ")
    #    field_merge_field = builder.insert_field(aw.fields.FieldType.FIELD_MERGE_FIELD, True).as_field_merge_field()
    #    field_merge_field.field_name = "Name"
    #    builder.writeln(",")

    #    # A MERGEREC field will print the row number of the data being merged in every merge output document.
    #    builder.write("\nRow number of record in data source: ")
    #    field_merge_rec = builder.insert_field(aw.fields.FieldType.FIELD_MERGE_REC, True).as_field_merge_rec()

    #    self.assertEqual(" MERGEREC ", field_merge_rec.get_field_code())

    #    # A MERGESEQ field will count the number of successful merges and print the current value on each respective page.
    #    # If a mail merge skips no rows and invokes no SKIP/SKIPIF/NEXT/NEXTIF fields, then all merges are successful.
    #    # The MERGESEQ and MERGEREC fields will display the same results of their mail merge was successful.
    #    builder.write("\nSuccessful merge number: ")
    #    field_merge_seq = builder.insert_field(aw.fields.FieldType.FIELD_MERGE_SEQ, True).as_field_merge_seq()

    #    self.assertEqual(" MERGESEQ ", field_merge_seq.get_field_code())

    #    # Insert a SKIPIF field, which will skip a merge if the name is "John Doe".
    #    field_skip_if = builder.insert_field(aw.fields.FieldType.FIELD_SKIP_IF, True).as_field_skip_if()
    #    builder.move_to(field_skip_if.separator)
    #    field_merge_field = builder.insert_field(aw.fields.FieldType.FIELD_MERGE_FIELD, True).as_field_merge_field()
    #    field_merge_field.field_name = "Name"
    #    field_skip_if.left_expression = "="
    #    field_skip_if.right_expression = "John Doe"

    #    # Create a data source with 3 rows, one of them having "John Doe" as a value for the "Name" column.
    #    # Since a SKIPIF field will be triggered once by that value, the output of our mail merge will have 2 pages instead of 3.
    #    # On page 1, the MERGESEQ and MERGEREC fields will both display "1".
    #    # On page 2, the MERGEREC field will display "3" and the MERGESEQ field will display "2".
    #    table = DataTable("Employees")
    #    table.columns.add("Name")
    #    table.rows.add(["Jane Doe"])
    #    table.rows.add(["John Doe"])
    #    table.rows.add(["Joe Bloggs"])

    #    doc.mail_merge.execute(table)
    #    doc.save(ARTIFACTS_DIR + "Field.mergerec.mergeseq.docx")
    #    #ExEnd

    #    doc = aw.Document(ARTIFACTS_DIR + "Field.mergerec.mergeseq.docx")

    #    self.assertEqual(0, doc.range.fields.count)

    #    self.assertEqual("Dear Jane Doe,\r" +
    #                    "\r" +
    #                    "Row number of record in data source: 1\r" +
    #                    "Successful merge number: 1\fDear Joe Bloggs,\r" +
    #                    "\r" +
    #                    "Row number of record in data source: 2\r" +
    #                    "Successful merge number: 3", doc.get_text().strip())

    def test_field_ocx(self):

        #ExStart
        #ExFor:FieldOcx
        #ExSummary:Shows how to insert an OCX field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        field = builder.insert_field(aw.fields.FieldType.FIELD_OCX, True).as_field_ocx()

        self.assertEqual(" OCX ", field.get_field_code())
        #ExEnd

        self.verify_field(aw.fields.FieldType.FIELD_OCX, " OCX ", "", field)

    ##ExStart
    ##ExFor:Field.remove
    ##ExFor:FieldPrivate
    ##ExSummary:Shows how to process PRIVATE fields.
    #def test_field_private(self):

    #    # Open a Corel WordPerfect document which we have converted to .docx format.
    #    doc = aw.Document(MY_DIR + "Field sample - PRIVATE.docx")

    #    # WordPerfect 5.x/6.x documents like the one we have loaded may contain PRIVATE fields.
    #    # Microsoft Word preserves PRIVATE fields during load/save operations,
    #    # but provides no functionality for them.
    #    field = doc.range.fields[0].as_field_private()

    #    self.assertEqual(" PRIVATE \"My value\" ", field.get_field_code())
    #    self.assertEqual(aw.fields.FieldType.FIELD_PRIVATE, field.type)

    #    # We can also insert PRIVATE fields using a document builder.
    #    builder = aw.DocumentBuilder(doc)
    #    builder.insert_field(aw.fields.FieldType.FIELD_PRIVATE, True)

    #    # These fields are not a viable way of protecting sensitive information.
    #    # Unless backward compatibility with older versions of WordPerfect is essential,
    #    # we can safely remove these fields. We can do this using a DocumentVisiitor implementation.
    #    self.assertEqual(2, doc.range.fields.count)

    #    remover = FieldPrivateRemover()
    #    doc.accept(remover)

    #    self.assertEqual(2, remover.get_fields_removed_count())
    #    self.assertEqual(0, doc.range.fields.count)

    #class FieldPrivateRemover(aw.DocumentVisitor):
    #    """Removes all encountered PRIVATE fields."""

    #    def __init__(self):

    #        self.fields_removed_count = 0

    #    def get_fields_removed_count(self) -> int:

    #        return self.fields_removed_count

    #    def visit_field_end(self, field_end: aw.fields.FieldEnd) -> aw.VisitorAction:
    #        """Called when a FieldEnd node is encountered in the document.
    #        If the node belongs to a PRIVATE field, the entire field is removed."""

    #        if field_end.field_type == aw.fields.FieldType.FIELD_PRIVATE:
    #            field_end.get_field().remove()
    #            self.fields_removed_count += 1

    #        return aw.VisitorAction.CONTINUE

    ##ExEnd

    def test_field_section(self):

        #ExStart
        #ExFor:FieldSection
        #ExFor:FieldSectionPages
        #ExSummary:Shows how to use SECTION and SECTIONPAGES fields to number pages by sections.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.paragraph_format.alignment = aw.ParagraphAlignment.RIGHT

        # A SECTION field displays the number of the section it is in.
        builder.write("Section ")
        field_section = builder.insert_field(aw.fields.FieldType.FIELD_SECTION, True).as_field_section()

        self.assertEqual(" SECTION ", field_section.get_field_code())

        # A PAGE field displays the number of the page it is in.
        builder.write("\nPage ")
        field_page = builder.insert_field(aw.fields.FieldType.FIELD_PAGE, True).as_field_page()

        self.assertEqual(" PAGE ", field_page.get_field_code())

        # A SECTIONPAGES field displays the number of pages that the section it is in spans across.
        builder.write(" of ")
        field_section_pages = builder.insert_field(aw.fields.FieldType.FIELD_SECTION_PAGES, True).as_field_section_pages()

        self.assertEqual(" SECTIONPAGES ", field_section_pages.get_field_code())

        # Move out of the header back into the main document and insert two pages.
        # All these pages will be in the first section. Our fields, which appear once every header,
        # will number the current/total pages of this section.
        builder.move_to_document_end()
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # We can insert a new section with the document builder like this.
        # This will affect the values displayed in the SECTION and SECTIONPAGES fields in all upcoming headers.
        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)

        # The PAGE field will keep counting pages across the whole document.
        # We can manually reset its count at each section to keep track of pages section-by-section.
        builder.current_section.page_setup.restart_page_numbering = True
        builder.insert_break(aw.BreakType.PAGE_BREAK)

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "Field.field_section.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_section.docx")

        self.verify_field(aw.fields.FieldType.FIELD_SECTION, " SECTION ", "2", doc.range.fields[0])
        self.verify_field(aw.fields.FieldType.FIELD_PAGE, " PAGE ", "2", doc.range.fields[1])
        self.verify_field(aw.fields.FieldType.FIELD_SECTION_PAGES, " SECTIONPAGES ", "2", doc.range.fields[2])

    #ExStart
    #ExFor:FieldTime
    #ExSummary:Shows how to display the current time using the TIME field.
    def test_field_time(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # By default, time is displayed in the "h:mm am/pm" format.
        field = ExField.insert_field_time(builder, "")

        self.assertEqual(" TIME ", field.get_field_code())

        # We can use the \@ flag to change the format of our displayed time.
        field = ExField.insert_field_time(builder, "\\@ HHmm")

        self.assertEqual(" TIME \\@ HHmm", field.get_field_code())

        # We can adjust the format to get TIME field to also display the date, according to the Gregorian calendar.
        field = ExField.insert_field_time(builder, "\\@ \"M/d/yyyy h mm:ss am/pm\"")

        self.assertEqual(" TIME \\@ \"M/d/yyyy h mm:ss am/pm\"", field.get_field_code())

        doc.save(ARTIFACTS_DIR + "Field.field_time.docx")
        self._test_field_time(aw.Document(ARTIFACTS_DIR + "Field.field_time.docx")) #ExSkip

    @staticmethod
    def insert_field_time(builder: aw.DocumentBuilder, format: str) -> aw.fields.FieldTime:
        """Use a document builder to insert a TIME field, insert a new paragraph and return the field."""

        field = builder.insert_field(aw.fields.FieldType.FIELD_TIME, True).as_field_time()
        builder.move_to(field.separator)
        builder.write(format)
        builder.move_to(field.start.parent_node)

        builder.insert_paragraph()
        return field

    #ExEnd

    def _test_field_time(self, doc: aw.Document):

        doc_loading_time = datetime.now()
        doc = DocumentHelper.save_open(doc)

        field = doc.range.fields[0].as_field_time()

        self.assertEqual(" TIME ", field.get_field_code())
        self.assertEqual(aw.fields.FieldType.FIELD_TIME, field.type)
        self.assertEqual(field.result, doc_loading_time.strftime("%I:%M %p").lower().lstrip('0'))

        field = doc.range.fields[1].as_field_time()

        self.assertEqual(" TIME \\@ HHmm", field.get_field_code())
        self.assertEqual(aw.fields.FieldType.FIELD_TIME, field.type)
        self.assertEqual(field.result, doc_loading_time.strftime("%I:%M %p").lower().lstrip('0'))

        field = doc.range.fields[2].as_field_time()

        self.assertEqual(" TIME \\@ \"M/d/yyyy h mm:ss am/pm\"", field.get_field_code())
        self.assertEqual(aw.fields.FieldType.FIELD_TIME, field.type)
        self.assertEqual(field.result, doc_loading_time.strftime("%I:%M %p").lower().lstrip('0'))

    def test_field_bidi_outline(self):

        #ExStart
        #ExFor:FieldBidiOutline
        #ExFor:FieldShape
        #ExFor:FieldShape.text
        #ExFor:ParagraphFormat.bidi
        #ExSummary:Shows how to create right-to-left language-compatible lists with BIDIOUTLINE fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # The BIDIOUTLINE field numbers paragraphs like the AUTONUM/LISTNUM fields,
        # but is only visible when a right-to-left editing language is enabled, such as Hebrew or Arabic.
        # The following field will display ".1", the RTL equivalent of list number "1.".
        field = builder.insert_field(aw.fields.FieldType.FIELD_BIDI_OUTLINE, True).as_field_bidi_outline()
        builder.writeln("שלום")

        self.assertEqual(" BIDIOUTLINE ", field.get_field_code())

        # Add two more BIDIOUTLINE fields, which will display ".2" and ".3".
        builder.insert_field(aw.fields.FieldType.FIELD_BIDI_OUTLINE, True)
        builder.writeln("שלום")
        builder.insert_field(aw.fields.FieldType.FIELD_BIDI_OUTLINE, True)
        builder.writeln("שלום")

        # Set the horizontal text alignment for every paragraph in the document to RTL.
        for para in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True):
            para = para.as_paragraph()
            para.paragraph_format.bidi = True

        # If we enable a right-to-left editing language in Microsoft Word, our fields will display numbers.
        # Otherwise, they will display "###".
        doc.save(ARTIFACTS_DIR + "Field.field_bidi_outline.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Field.field_bidi_outline.docx")

        for field_bidi_outline in doc.range.fields:
            self.verify_field(aw.fields.FieldType.FIELD_BIDI_OUTLINE, " BIDIOUTLINE ", "", field_bidi_outline)

    def test_legacy(self):

        #ExStart
        #ExFor:FieldEmbed
        #ExFor:FieldShape
        #ExFor:FieldShape.text
        #ExSummary:Shows how some older Microsoft Word fields such as SHAPE and EMBED are handled during loading.
        # Open a document that was created in Microsoft Word 2003.
        doc = aw.Document(MY_DIR + "Legacy fields.doc")

        # If we open the Word document and press Alt+F9, we will see a SHAPE and an EMBED field.
        # A SHAPE field is the anchor/canvas for an AutoShape object with the "In line with text" wrapping style enabled.
        # An EMBED field has the same function, but for an embedded object,
        # such as a spreadsheet from an external Excel document.
        # However, these fields will not appear in the document's Fields collection.
        self.assertEqual(0, doc.range.fields.count)

        # These fields are supported only by old versions of Microsoft Word.
        # The document loading process will convert these fields into Shape objects,
        # which we can access in the document's node collection.
        shapes = doc.get_child_nodes(aw.NodeType.SHAPE, True)
        self.assertEqual(3, shapes.count)

        # The first Shape node corresponds to the SHAPE field in the input document,
        # which is the inline canvas for the AutoShape.
        shape = shapes[0].as_shape()
        self.assertEqual(aw.drawing.ShapeType.IMAGE, shape.shape_type)

        # The second Shape node is the AutoShape itself.
        shape = shapes[1].as_shape()
        self.assertEqual(aw.drawing.ShapeType.CAN, shape.shape_type)

        # The third Shape is what was the EMBED field that contained the external spreadsheet.
        shape = shapes[2].as_shape()
        self.assertEqual(aw.drawing.ShapeType.OLE_OBJECT, shape.shape_type)
        #ExEnd

    def test_set_field_index_format(self):

        #ExStart
        #ExFor:FieldOptions.field_index_format
        #ExSummary:Shows how to formatting FieldIndex fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.write("A")
        builder.insert_break(aw.BreakType.LINE_BREAK)
        builder.insert_field("XE \"A\"")
        builder.write("B")

        builder.insert_field(" INDEX \\e \" · \" \\h \"A\" \\c \"2\" \\z \"1033\"", None)

        doc.field_options.field_index_format = aw.fields.FieldIndexFormat.FANCY
        doc.update_fields()

        doc.save(ARTIFACTS_DIR + "Field.set_field_index_format.docx")
        #ExEnd

    ##ExStart
    ##ExFor:ComparisonEvaluationResult.__init__(bool)
    ##ExFor:ComparisonEvaluationResult.__init__(string)
    ##ExFor:ComparisonEvaluationResult
    ##ExFor:ComparisonExpression
    ##ExFor:ComparisonExpression.left_expression
    ##ExFor:ComparisonExpression.comparison_operator
    ##ExFor:ComparisonExpression.right_expression
    ##ExFor:FieldOptions.comparison_expression_evaluator
    ##ExSummary:Shows how to implement custom evaluation for the IF and COMPARE fields.
    #def test_condition_evaluation_extension_point(self):

    #    parameters = [
    #        (" IF {0} {1} {2} \"True argument\" \"False argument\" ", 1, None, "True argument"),
    #        (" IF {0} {1} {2} \"True argument\" \"False argument\" ", 0, None, "False argument"),
    #        (" IF {0} {1} {2} \"True argument\" \"False argument\" ", -1, "Custom Error", "Custom Error"),
    #        (" IF {0} {1} {2} \"True argument\" \"False argument\" ", -1, None, "True argument"),
    #        (" COMPARE {0} {1} {2} ", 1, None, "1"),
    #        (" COMPARE {0} {1} {2} ", 0, None, "0"),
    #        (" COMPARE {0} {1} {2} ", -1, "Custom Error", "Custom Error"),
    #        (" COMPARE {0} {1} {2} ", -1, None, "1")]

    #    for field_code, comparison_result, comparison_error, expected_result in parameters:
    #        with self.subTest(field_code=field_code,
    #                          comparison_result=comparison_result,
    #                          comparison_error=comparison_error,
    #                          expected_result=expected_result):
    #            left = '"left expression"'
    #            operator = '<>'
    #            right = '"right expression"'

    #            builder = aw.DocumentBuilder()

    #            # Field codes that we use in this example:
    #            # 1.   " IF {0} {1} {2} \"true argument\" \"false argument\" ".
    #            # 2.   " COMPARE {0} {1} {2} ".
    #            field = builder.insert_field(field_code.format(left, operator, right), None)

    #            # If the "comparison_result" is undefined, we create "ComparisonEvaluationResult" with string, instead of bool.
    #            if comparison_result != -1:
    #                result = ComparisonEvaluationResult(comparison_result == 1)
    #            elif comparison_error is not None:
    #                result = ComparisonEvaluationResult(comparison_error)
    #            else:
    #                result = None

    #            evaluator = ExField.ComparisonExpressionEvaluator(result)
    #            builder.document.field_options.comparison_expression_evaluator = evaluator

    #            builder.document.update_fields()

    #            self.assertEqual(expected_result, field.result)
    #            evaluator.assert_invocations_count(1).assert_invocation_arguments(0, left, operator, right)

    #class ComparisonExpressionEvaluator(aw.IComparisonExpressionEvaluator):
    #    """Comparison expressions evaluation for the FieldIf and FieldCompare."""

    #    def __init__(self, result: ComparisonEvaluationResult):

    #        self.result = result
    #        self.invocations: List[List[str]] = []

    #    def evaluate(self, field: aw.fields.Field, expression: ComparisonExpression) -> ComparisonEvaluationResult:

    #        self.invocations.add([
    #            expression.left_expression,
    #            expression.comparison_operator,
    #            expression.right_expression
    #        ])

    #        return self.result

    #    def assert_invocations_count(self, expected: int) -> ComparisonExpressionEvaluator:

    #        self.assertEqual(expected, self.invocations.count)
    #        return self

    #    def assert_invocation_arguments(
    #        invocation_index: int,
    #        expected_left_expression: str,
    #        expected_comparison_operator: str,
    #        expected_right_expression: str) -> ComparisonExpressionEvaluator:

    #        arguments = self.invocations[invocation_index]

    #        self.assertEqual(expected_left_expression, arguments[0])
    #        self.assertEqual(expected_comparison_operator, arguments[1])
    #        self.assertEqual(expected_right_expression, arguments[2])

    #        return self

    ##ExEnd

    #def test_comparison_expression_evaluator_nested_fields(self):

    #    document = aw.Document()

    #    field_builder = aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_IF)
    #    field_builder.add_argument(
    #        aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_IF)
    #            .add_argument(123)
    #            .add_argument(">")
    #            .add_argument(666)
    #            .add_argument("left greater than right")
    #            .add_argument("left less than right"))
    #    field_builder.add_argument("<>")
    #    field_builder.add_argument(
    #        aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_IF)
    #            .add_argument("left expression")
    #            .add_argument("=")
    #            .add_argument("right expression")
    #            .add_argument("expression are equal")
    #            .add_argument("expression are not equal"))
    #    field_builder.add_argument(
    #        aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_IF)
    #            .add_argument(aw.fields.FieldArgumentBuilder()
    #                .add_text("#")
    #                .add_field(aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_PAGE)))
    #            .add_argument("=")
    #            .add_argument(aw.fields.FieldArgumentBuilder()
    #                .add_text("#")
    #                .add_field(aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_NUM_PAGES)))
    #            .add_argument("the last page")
    #            .add_argument("not the last page"))
    #    field_builder.add_argument(
    #        aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_IF)
    #            .add_argument("unexpected")
    #            .add_argument("=")
    #            .add_argument("unexpected")
    #            .add_argument("unexpected")
    #            .add_argument("unexpected"))
    #    field_builder.build_and_insert(document.first_section.body.first_paragraph)

    #    evaluator = ExField.ComparisonExpressionEvaluator(None)
    #    document.field_options.comparison_expression_evaluator = evaluator

    #    document.update_fields()

    #    evaluator.assert_invocations_count(4)
    #    evaluator.assert_invocation_arguments(0, "123", ">", "666")
    #    evaluator.assert_invocation_arguments(1, "\"left expression\"", "=", "\"right expression\"")
    #    evaluator.assert_invocation_arguments(2, "left less than right", "<>", "expression are not equal")
    #    evaluator.assert_invocation_arguments(3, "\"#1\"", "=", "\"#1\"")

    #def test_comparison_expression_evaluator_header_footer_fields(self):

    #    document = aw.Document()
    #    builder = aw.DocumentBuilder(document)

    #    builder.insert_break(aw.BreakType.PAGE_BREAK)
    #    builder.insert_break(aw.BreakType.PAGE_BREAK)
    #    builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)

    #    field_builder = aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_IF)
    #    field_builder.add_argument(aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_PAGE))
    #    field_builder.add_argument("=")
    #    field_builder.add_argument(aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_NUM_PAGES))
    #    field_builder.add_argument(aw.fields.FieldArgumentBuilder()
    #        .add_field(aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_PAGE))
    #        .add_text(" / ")
    #        .add_field(aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_NUM_PAGES)))
    #    field_builder.add_argument(aw.fields.FieldArgumentBuilder()
    #        .add_field(aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_PAGE))
    #        .add_text(" / ")
    #        .add_field(aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_NUM_PAGES)))
    #    field_builder.build_and_insert(builder.current_paragraph)

    #    evaluator = ExField.ComparisonExpressionEvaluator(None)
    #    document.field_options.comparison_expression_evaluator = evaluator

    #    document.update_fields()

    #    evaluator.assert_invocations_count(3)
    #    evaluator.assert_invocation_arguments(0, "1", "=", "3")
    #    evaluator.assert_invocation_arguments(1, "2", "=", "3")
    #    evaluator.assert_invocation_arguments(2, "3", "=", "3")

    ##ExStart
    ##ExFor:IFieldUpdatingCallback
    ##ExFor:IFieldUpdatingCallback.field_updating(Field)
    ##ExFor:IFieldUpdatingCallback.field_updated(Field)
    ##ExSummary:Shows how to use callback methods during a field update.
    #def test_field_updating_callback_test(self):

    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    builder.insert_field(" DATE \\@ \"dddd, d MMMM yyyy\" ")
    #    builder.insert_field(" TIME ")
    #    builder.insert_field(" REVNUM ")
    #    builder.insert_field(" AUTHOR  \"John Doe\" ")
    #    builder.insert_field(" SUBJECT \"My Subject\" ")
    #    builder.insert_field(" QUOTE \"Hello world!\" ")

    #    callback = ExField.FieldUpdatingCallback()
    #    doc.field_options.field_updating_callback = callback

    #    doc.update_fields()

    #    self.assertTrue(callback.field_updated_calls.contains("Updating John Doe"))

    #class FieldUpdatingCallback(aw.fields.IFieldUpdatingCallback):
    #    """Implement this interface if you want to have your own custom methods called during a field update."""

    #    def __init__(self):

    #        self.field_updated_calls: List[str] = []

    #    def field_updating(self, field: aw.fields.Field):
    #        """A user defined method that is called just before a field is updated."""

    #        if field.type == aw.fields.FieldType.FIELD_AUTHOR:
    #            field_author = field.as_field_author()
    #            field_author.author_name = "Updating John Doe"

    #    def field_updated(self, field: aw.fields.Field):
    #        """A user defined method that is called just after a field is updated."""

    #        self.field_updated_calls.add(field.result)

    ##ExEnd
