import unittest
import io

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir
from document_helper import DocumentHelper
from testutil import TestUtil

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

class ExFieldOptions(ApiExampleBase):

    def test_current_user(self):

        #ExStart
        #ExFor:Document.UpdateFields
        #ExFor:FieldOptions.CurrentUser
        #ExFor:UserInformation
        #ExFor:UserInformation.Name
        #ExFor:UserInformation.Initials
        #ExFor:UserInformation.Address
        #ExFor:UserInformation.DefaultUser
        #ExSummary:Shows how to set user details, and display them using fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create a UserInformation object and set it as the data source for fields that display user information.
        user_information = aw.fields.UserInformation()
        user_information.name = "John Doe"
        user_information.initials = "J. D."
        user_information.address = "123 Main Street"

        doc.field_options.current_user = user_information

        # Insert USERNAME, USERINITIALS, and USERADDRESS fields, which display values of
        # the respective properties of the UserInformation object that we have created above.
        self.assertEqual(user_information.name, builder.insert_field(" USERNAME ").result)
        self.assertEqual(user_information.initials, builder.insert_field(" USERINITIALS ").result)
        self.assertEqual(user_information.address, builder.insert_field(" USERADDRESS ").result)

        # The field options object also has a static default user that fields from all documents can refer to.
        user_information.default_user.name = "Default User"
        user_information.default_user.initials = "D. U."
        user_information.default_user.address = "One Microsoft Way"
        doc.field_options.current_user = aw.fields.UserInformation.default_user

        self.assertEqual("Default User", builder.insert_field(" USERNAME ").result)
        self.assertEqual("D. U.", builder.insert_field(" USERINITIALS ").result)
        self.assertEqual("One Microsoft Way", builder.insert_field(" USERADDRESS ").result)

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "FieldOptions.current_user.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "FieldOptions.current_user.docx")

        self.assertIsNone(doc.field_options.current_user)

        field_user_name = doc.range.fields[0].as_field_user_name()

        self.assertIsNone(field_user_name.user_name)
        self.assertEqual("Default User", field_user_name.result)

        field_user_initials = doc.range.fields[1].as_field_user_initials()

        self.assertIsNone(field_user_initials.user_initials)
        self.assertEqual("D. U.", field_user_initials.result)

        field_user_address = doc.range.fields[2].as_field_user_address()

        self.assertIsNone(field_user_address.user_address)
        self.assertEqual("One Microsoft Way", field_user_address.result)

    def test_file_name(self):

        #ExStart
        #ExFor:FieldOptions.FileName
        #ExFor:FieldFileName
        #ExFor:FieldFileName.IncludeFullPath
        #ExSummary:Shows how to use FieldOptions to override the default value for the FILENAME field.
        doc = aw.Document(MY_DIR + "Document.docx")
        builder = aw.DocumentBuilder(doc)

        builder.move_to_document_end()
        builder.writeln()

        # This FILENAME field will display the local system file name of the document we loaded.
        field = builder.insert_field(aw.fields.FieldType.FIELD_FILE_NAME, True).as_field_file_name()
        field.update()

        self.assertEqual(" FILENAME ", field.get_field_code())
        self.assertEqual("Document.docx", field.result)

        builder.writeln()

        # By default, the FILENAME field shows the file's name, but not its full local file system path.
        # We can set a flag to make it show the full file path.
        field = builder.insert_field(aw.fields.FieldType.FIELD_FILE_NAME, True).as_field_file_name()
        field.include_full_path = True
        field.update()

        self.assertEqual(MY_DIR + "Document.docx", field.result)

        # We can also set a value for this property to
        # override the value that the FILENAME field displays.
        doc.field_options.file_name = "FieldOptions.filename.docx"
        field.update()

        self.assertEqual(" FILENAME  \\p", field.get_field_code())
        self.assertEqual("FieldOptions.filename.docx", field.result)

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + doc.field_options.file_name)
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "FieldOptions.filename.docx")

        self.assertIsNone(doc.field_options.file_name)
        TestUtil.verify_field(self, aw.fields.FieldType.FIELD_FILE_NAME, " FILENAME ", "FieldOptions.filename.docx", doc.range.fields[0])

    def test_bidi(self):

        #ExStart
        #ExFor:FieldOptions.IsBidiTextSupportedOnUpdate
        #ExSummary:Shows how to use FieldOptions to ensure that field updating fully supports bi-directional text.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Ensure that any field operation involving right-to-left text is performs as expected.
        doc.field_options.is_bidi_text_supported_on_update = True

        # Use a document builder to insert a field that contains the right-to-left text.
        combo_box = builder.insert_combo_box("MyComboBox", ["עֶשְׂרִים", "שְׁלוֹשִׁים", "אַרְבָּעִים", "חֲמִשִּׁים", "שִׁשִּׁים"], 0)
        combo_box.calculate_on_exit = True

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "FieldOptions.bidi.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "FieldOptions.bidi.docx")

        self.assertFalse(doc.field_options.is_bidi_text_supported_on_update)

        combo_box = doc.range.form_fields[0]

        self.assertEqual("עֶשְׂרִים", combo_box.result)

    def test_legacy_number_format(self):

        #ExStart
        #ExFor:FieldOptions.LegacyNumberFormat
        #ExSummary:Shows how enable legacy number formatting for fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        field = builder.insert_field("= 2 + 3 \\# $##")

        self.assertEqual("$ 5", field.result)

        doc.field_options.legacy_number_format = True
        field.update()

        self.assertEqual("$5", field.result)
        #ExEnd

        doc = DocumentHelper.save_open(doc)

        self.assertFalse(doc.field_options.legacy_number_format)
        TestUtil.verify_field(self, aw.fields.FieldType.FIELD_FORMULA, "= 2 + 3 \\# $##", "$5", doc.range.fields[0])

    def test_pre_process_culture(self):

        #ExStart
        #ExFor:FieldOptions.PreProcessCulture
        #ExSummary:Shows how to set the preprocess culture.
        doc = aw.Document(MY_DIR + "Document.docx")
        builder = aw.DocumentBuilder(doc)

        # Set the culture according to which some fields will format their displayed values.
        doc.field_options.pre_process_culture = "de-DE"

        field = builder.insert_field(" DOCPROPERTY CreateTime")

        # The DOCPROPERTY field will display its result formatted according to the preprocess culture
        # we have set to German. The field will display the date/time using the "dd.mm.yyyy hh:mm" format.
        self.assertRegex(field.result, r"\d{2}[.]\d{2}[.]\d{4} \d{2}[:]\d{2}")

        doc.field_options.pre_process_culture = ""
        field.update()

        # After switching to the invariant culture, the DOCPROPERTY field will use the "mm/dd/yyyy hh:mm" format.
        self.assertRegex(field.result, r"\d{2}[/]\d{2}[/]\d{4} \d{2}[:]\d{2}")
        #ExEnd

        doc = DocumentHelper.save_open(doc)

        self.assertIsNone(doc.field_options.pre_process_culture)
        self.assertRegex(doc.range.fields[0].result, r"\d{2}[/]\d{2}[/]\d{4} \d{2}[:]\d{2}")

    def test_table_of_authority_categories(self):

        #ExStart
        #ExFor:FieldOptions.ToaCategories
        #ExFor:ToaCategories
        #ExFor:ToaCategories.Item(Int32)
        #ExFor:ToaCategories.DefaultCategories
        #ExSummary:Shows how to specify a set of categories for TOA fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # TOA fields can filter their entries by categories defined in this collection.
        toa_categories = aw.fields.ToaCategories()
        doc.field_options.toa_categories = toa_categories

        # This collection of categories comes with default values, which we can overwrite with custom values.
        self.assertEqual("Cases", toa_categories[1])
        self.assertEqual("Statutes", toa_categories[2])

        toa_categories[1] = "My Category 1"
        toa_categories[2] = "My Category 2"

        # We can always access the default values via this collection.
        self.assertEqual("Cases", aw.fields.ToaCategories.default_categories[1])
        self.assertEqual("Statutes", aw.fields.ToaCategories.default_categories[2])

        # Insert 2 TOA fields. TOA fields create an entry for each TA field in the document.
        # Use the "\c" switch to select the index of a category from our collection.
        #  With this switch, a TOA field will only pick up entries from TA fields that
        # also have a "\c" switch with a matching category index. Each TOA field will also display
        # the name of the category that its "\c" switch points to.
        builder.insert_field("TOA \\c 1 \\h", None)
        builder.insert_field("TOA \\c 2 \\h", None)
        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # Insert TOA entries across 2 categories. Our first TOA field will receive one entry,
        # from the second TA field whose "\c" switch also points to the first category.
        # The second TOA field will have two entries from the other two TA fields.
        builder.insert_field("TA \\c 2 \\l \"entry 1\"")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.insert_field("TA \\c 1 \\l \"entry 2\"")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.insert_field("TA \\c 2 \\l \"entry 3\"")

        doc.update_fields()
        doc.save(ARTIFACTS_DIR + "FieldOptions.t_o_a.categories.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "FieldOptions.t_o_a.categories.docx")

        self.assertIsNone(doc.field_options.toa_categories)

        TestUtil.verify_field(self, aw.fields.FieldType.FIELD_TOA, "TOA \\c 1 \\h", "My Category 1\rentry 2\t3\r", doc.range.fields[0])
        TestUtil.verify_field(self, aw.fields.FieldType.FIELD_TOA, "TOA \\c 2 \\h",
            "My Category 2\r" +
            "entry 1\t2\r" +
            "entry 3\t4\r", doc.range.fields[1])

    def test_use_invariant_culture_number_format(self):

        #ExStart
        #ExFor:FieldOptions.UseInvariantCultureNumberFormat
        #ExSummary:Shows how to format numbers according to the invariant culture.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        Thread.current_thread.current_culture = "de-DE"
        field = builder.insert_field(" = 1234567,89 \\# $#,###,###.##")
        field.update()

        # Sometimes, fields may not format their numbers correctly under certain cultures.
        self.assertFalse(doc.field_options.use_invariant_culture_number_format)
        self.assertEqual("$1234567,89 .     ", field.result)

        # To fix this, we could change the culture for the entire thread.
        # Another way to fix this is to set this flag,
        # which gets all fields to use the invariant culture when formatting numbers.
        # This way allows us to avoid changing the culture for the entire thread.
        doc.field_options.use_invariant_culture_number_format = True
        field.update()
        self.assertEqual("$1.234.567,89", field.result)
        #ExEnd

        doc = DocumentHelper.save_open(doc)

        self.assertFalse(doc.field_options.use_invariant_culture_number_format)
        TestUtil.verify_field(self, aw.fields.FieldType.FIELD_FORMULA, " = 1234567,89 \\# $#,###,###.##", "$1.234.567,89", doc.range.fields[0])

    ##ExStart
    ##ExFor:FieldOptions.FieldUpdateCultureProvider
    ##ExFor:IFieldUpdateCultureProvider
    ##ExFor:IFieldUpdateCultureProvider.GetCulture(string, Field)
    ##ExSummary:Shows how to specify a culture which parses date/time formatting for each field.

    #def test_define_date_time_formatting(self):

    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    builder.insert_field(aw.fields.FieldType.FIELD_TIME, True)

    #    doc.field_options.field_update_culture_source = aw.fields.FieldUpdateCultureSource.FIELD_CODE

    #    # Set a provider that returns a culture object specific to each field.
    #    doc.field_options.field_update_culture_provider = ExFieldOptions.FieldUpdateCultureProvider()

    #    field_date = doc.range.fields[0].as_field_time()
    #    if field_date.locale_id != aw.loading.EditingLanguage.RUSSIAN:
    #        field_date.locale_id = aw.loading.EditingLanguage.RUSSIAN

    #    doc.save(ARTIFACTS_DIR + "FieldOptions.update_date_time_formatting.pdf")

    #class FieldUpdateCultureProvider(aw.fields.IFieldUpdateCultureProvider):
    #    """Provides a CultureInfo object that should be used during the update of a field."""

    #    def get_culture(self, name: str, field: aw.fields.Field) -> str:
    #        """Returns a CultureInfo object to be used during the field's update."""

    #        if name == "ru-RU":
    #            culture = CultureInfo(name, False)
    #            format = culture.date_time_format

    #            format.month_names = ["месяц 1", "месяц 2", "месяц 3", "месяц 4", "месяц 5", "месяц 6", "месяц 7", "месяц 8", "месяц 9", "месяц 10", "месяц 11", "месяц 12", ""]
    #            format.month_genitive_names = format.month_names
    #            format.abbreviated_month_names = ["мес 1", "мес 2", "мес 3", "мес 4", "мес 5", "мес 6", "мес 7", "мес 8", "мес 9", "мес 10", "мес 11", "мес 12", ""]
    #            format.abbreviated_month_genitive_names = format.abbreviated_month_names

    #            format.day_names = ["день недели 7", "день недели 1", "день недели 2", "день недели 3", "день недели 4", "день недели 5", "день недели 6"]
    #            format.abbreviated_day_names = ["день 7", "день 1", "день 2", "день 3", "день 4", "день 5", "день 6"]
    #            format.shortest_day_names = ["д7", "д1", "д2", "д3", "д4", "д5", "д6"]

    #            format.a_m_designator = "До полудня"
    #            format.p_m_designator = "После полудня"

    #            pattern = "yyyy MM (MMMM) dd (dddd) hh:mm:ss tt"
    #            format.long_date_pattern = pattern
    #            format.long_time_pattern = pattern
    #            format.short_date_pattern = pattern
    #            format.short_time_pattern = pattern

    #            return culture

    #        if name == "en-US":
    #            return name

    #        return None

    ##ExEnd

    #def test_barcode_generator(self):

    #    #ExStart
    #    #ExFor:BarcodeParameters
    #    #ExFor:BarcodeParameters.AddStartStopChar
    #    #ExFor:BarcodeParameters.BackgroundColor
    #    #ExFor:BarcodeParameters.BarcodeType
    #    #ExFor:BarcodeParameters.BarcodeValue
    #    #ExFor:BarcodeParameters.CaseCodeStyle
    #    #ExFor:BarcodeParameters.DisplayText
    #    #ExFor:BarcodeParameters.ErrorCorrectionLevel
    #    #ExFor:BarcodeParameters.FacingIdentificationMark
    #    #ExFor:BarcodeParameters.FixCheckDigit
    #    #ExFor:BarcodeParameters.ForegroundColor
    #    #ExFor:BarcodeParameters.IsBookmark
    #    #ExFor:BarcodeParameters.IsUSPostalAddress
    #    #ExFor:BarcodeParameters.PosCodeStyle
    #    #ExFor:BarcodeParameters.PostalAddress
    #    #ExFor:BarcodeParameters.ScalingFactor
    #    #ExFor:BarcodeParameters.SymbolHeight
    #    #ExFor:BarcodeParameters.SymbolRotation
    #    #ExFor:IBarcodeGenerator
    #    #ExFor:IBarcodeGenerator.GetBarcodeImage(BarcodeParameters)
    #    #ExFor:IBarcodeGenerator.GetOldBarcodeImage(BarcodeParameters)
    #    #ExFor:FieldOptions.BarcodeGenerator
    #    #ExSummary:Shows how to use a barcode generator.
    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)
    #    self.assertIsNone(doc.field_options.barcode_generator) #ExSkip

    #    # We can use a custom IBarcodeGenerator implementation to generate barcodes,
    #    # and then insert them into the document as images.
    #    doc.field_options.barcode_generator = CustomBarcodeGenerator()

    #    # Below are four examples of different barcode types that we can create using our generator.
    #    # For each barcode, we specify a new set of barcode parameters, and then generate the image.
    #    # Afterwards, we can insert the image into the document, or save it to the local file system.
    #    # 1 -  QR code:
    #    barcode_parameters = aw.fields.BarcodeParameters()
    #    barcode_parameters.barcode_type = "QR"
    #    barcode_parameters.barcode_value = "ABC123"
    #    barcode_parameters.background_color = "0xF8BD69"
    #    barcode_parameters.foreground_color = "0xB5413B"
    #    barcode_parameters.error_correction_level = "3"
    #    barcode_parameters.scaling_factor = "250"
    #    barcode_parameters.symbol_height = "1000"
    #    barcode_parameters.symbol_rotation = "0"

    #    img = doc.field_options.barcode_generator.get_barcode_image(barcode_parameters)
    #    img.save(ARTIFACTS_DIR + "FieldOptions.barcode_generator.qr.jpg")

    #    builder.insert_image(img)

    #    # 2 -  EAN13 barcode:
    #    barcode_parameters = aw.fields.BarcodeParameters()
    #    barcode_parameters.barcode_type = "EAN13"
    #    barcode_parameters.barcode_value = "501234567890"
    #    barcode_parameters.display_text = True
    #    barcode_parameters.pos_code_style = "CASE"
    #    barcode_parameters.fix_check_digit = True

    #    img = doc.field_options.barcode_generator.get_barcode_image(barcode_parameters)
    #    img.save(ARTIFACTS_DIR + "FieldOptions.barcode_generator.e_a_n13.jpg")
    #    builder.insert_image(img)

    #    # 3 -  CODE39 barcode:
    #    barcode_parameters = aw.fields.BarcodeParameters()
    #    barcode_parameters.barcode_type = "CODE39"
    #    barcode_parameters.barcode_value = "12345ABCDE"
    #    barcode_parameters.add_start_stop_char = True

    #    img = doc.field_options.barcode_generator.get_barcode_image(barcode_parameters)
    #    img.save(ARTIFACTS_DIR + "FieldOptions.barcode_generator.code39.jpg")
    #    builder.insert_image(img)

    #    # 4 -  ITF14 barcode:
    #    barcode_parameters = aw.fields.BarcodeParameters()
    #    barcode_parameters.barcode_type = "ITF14"
    #    barcode_parameters.barcode_value = "09312345678907"
    #    barcode_parameters.case_code_style = "STD"

    #    img = doc.field_options.barcode_generator.get_barcode_image(barcode_parameters)
    #    img.save(ARTIFACTS_DIR + "FieldOptions.barcode_generator.i_t_f14.jpg")
    #    builder.insert_image(img)

    #    doc.save(ARTIFACTS_DIR + "FieldOptions.barcode_generator.docx")
    #    #ExEnd

    #    TestUtil.verify_image(223, 223, ARTIFACTS_DIR + "FieldOptions.barcode_generator.q_r.jpg")
    #    TestUtil.verify_image(117, 108, ARTIFACTS_DIR + "FieldOptions.barcode_generator.e_a_n13.jpg")
    #    TestUtil.verify_image(397, 70, ARTIFACTS_DIR + "FieldOptions.barcode_generator.c_o_d_e39.jpg")
    #    TestUtil.verify_image(633, 134, ARTIFACTS_DIR + "FieldOptions.barcode_generator.i_t_f14.jpg")

    #    doc = aw.Document(ARTIFACTS_DIR + "FieldOptions.barcode_generator.docx")
    #    barcode = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

    #    self.assertTrue(barcode.has_image)
