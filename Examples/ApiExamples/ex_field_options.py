# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
from document_helper import DocumentHelper
import aspose.words as aw
import aspose.words.fields
import document_helper
import test_util
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExFieldOptions(ApiExampleBase):

    def test_current_user(self):
        #ExStart
        #ExFor:Document.update_fields
        #ExFor:FieldOptions.current_user
        #ExFor:UserInformation
        #ExFor:UserInformation.name
        #ExFor:UserInformation.initials
        #ExFor:UserInformation.address
        #ExFor:UserInformation.default_user
        #ExSummary:Shows how to set user details, and display them using fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Create a UserInformation object and set it as the data source for fields that display user information.
        user_information = aw.fields.UserInformation()
        user_information.name = 'John Doe'
        user_information.initials = 'J. D.'
        user_information.address = '123 Main Street'
        doc.field_options.current_user = user_information
        # Insert USERNAME, USERINITIALS, and USERADDRESS fields, which display values of
        # the respective properties of the UserInformation object that we have created above.
        self.assertEqual(user_information.name, builder.insert_field(field_code=' USERNAME ').result)
        self.assertEqual(user_information.initials, builder.insert_field(field_code=' USERINITIALS ').result)
        self.assertEqual(user_information.address, builder.insert_field(field_code=' USERADDRESS ').result)
        # The field options object also has a static default user that fields from all documents can refer to.
        aw.fields.UserInformation.default_user.name = 'Default User'
        aw.fields.UserInformation.default_user.initials = 'D. U.'
        aw.fields.UserInformation.default_user.address = 'One Microsoft Way'
        doc.field_options.current_user = aw.fields.UserInformation.default_user
        self.assertEqual('Default User', builder.insert_field(field_code=' USERNAME ').result)
        self.assertEqual('D. U.', builder.insert_field(field_code=' USERINITIALS ').result)
        self.assertEqual('One Microsoft Way', builder.insert_field(field_code=' USERADDRESS ').result)
        doc.update_fields()
        doc.save(file_name=ARTIFACTS_DIR + 'FieldOptions.CurrentUser.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'FieldOptions.CurrentUser.docx')
        self.assertIsNone(doc.field_options.current_user)
        field_user_name = doc.range.fields[0].as_field_user_name()
        self.assertIsNone(field_user_name.user_name)
        self.assertEqual('Default User', field_user_name.result)
        field_user_initials = doc.range.fields[1].as_field_user_initials()
        self.assertIsNone(field_user_initials.user_initials)
        self.assertEqual('D. U.', field_user_initials.result)
        field_user_address = doc.range.fields[2].as_field_user_address()
        self.assertIsNone(field_user_address.user_address)
        self.assertEqual('One Microsoft Way', field_user_address.result)

    def test_file_name(self):
        #ExStart
        #ExFor:FieldOptions.file_name
        #ExFor:FieldFileName
        #ExFor:FieldFileName.include_full_path
        #ExSummary:Shows how to use FieldOptions to override the default value for the FILENAME field.
        doc = aw.Document(file_name=MY_DIR + 'Document.docx')
        builder = aw.DocumentBuilder(doc=doc)
        builder.move_to_document_end()
        builder.writeln()
        # This FILENAME field will display the local system file name of the document we loaded.
        field = builder.insert_field(field_type=aw.fields.FieldType.FIELD_FILE_NAME, update_field=True).as_field_file_name()
        field.update()
        self.assertEqual(' FILENAME ', field.get_field_code())
        self.assertEqual('Document.docx', field.result)
        builder.writeln()
        # By default, the FILENAME field shows the file's name, but not its full local file system path.
        # We can set a flag to make it show the full file path.
        field = builder.insert_field(field_type=aw.fields.FieldType.FIELD_FILE_NAME, update_field=True).as_field_file_name()
        field.include_full_path = True
        field.update()
        self.assertEqual(MY_DIR + 'Document.docx', field.result)
        # We can also set a value for this property to
        # override the value that the FILENAME field displays.
        doc.field_options.file_name = 'FieldOptions.FILENAME.docx'
        field.update()
        self.assertEqual(' FILENAME  \\p', field.get_field_code())
        self.assertEqual('FieldOptions.FILENAME.docx', field.result)
        doc.update_fields()
        doc.save(file_name=ARTIFACTS_DIR + doc.field_options.file_name)
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'FieldOptions.FILENAME.docx')
        self.assertIsNone(doc.field_options.file_name)
        test_util.TestUtil.verify_field(expected_type=aw.fields.FieldType.FIELD_FILE_NAME, expected_field_code=' FILENAME ', expected_result='FieldOptions.FILENAME.docx', field=doc.range.fields[0])

    def test_bidi(self):
        #ExStart
        #ExFor:FieldOptions.is_bidi_text_supported_on_update
        #ExSummary:Shows how to use FieldOptions to ensure that field updating fully supports bi-directional text.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Ensure that any field operation involving right-to-left text is performs as expected.
        doc.field_options.is_bidi_text_supported_on_update = True
        # Use a document builder to insert a field that contains the right-to-left text.
        combo_box = builder.insert_combo_box('MyComboBox', ['עֶשְׂרִים', 'שְׁלוֹשִׁים', 'אַרְבָּעִים', 'חֲמִשִּׁים', 'שִׁשִּׁים'], 0)
        combo_box.calculate_on_exit = True
        doc.update_fields()
        doc.save(file_name=ARTIFACTS_DIR + 'FieldOptions.Bidi.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'FieldOptions.Bidi.docx')
        self.assertFalse(doc.field_options.is_bidi_text_supported_on_update)
        combo_box = doc.range.form_fields[0]
        self.assertEqual('עֶשְׂרִים', combo_box.result)

    def test_legacy_number_format(self):
        #ExStart
        #ExFor:FieldOptions.legacy_number_format
        #ExSummary:Shows how enable legacy number formatting for fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        field = builder.insert_field(field_code='= 2 + 3 \\# $##')
        self.assertEqual('$ 5', field.result)
        doc.field_options.legacy_number_format = True
        field.update()
        self.assertEqual('$5', field.result)
        #ExEnd
        doc = document_helper.DocumentHelper.save_open(doc)
        self.assertFalse(doc.field_options.legacy_number_format)
        test_util.TestUtil.verify_field(expected_type=aw.fields.FieldType.FIELD_FORMULA, expected_field_code='= 2 + 3 \\# $##', expected_result='$5', field=doc.range.fields[0])

    def test_table_of_authority_categories(self):
        #ExStart
        #ExFor:FieldOptions.toa_categories
        #ExFor:ToaCategories
        #ExFor:ToaCategories.__getitem__(int)
        #ExFor:ToaCategories.default_categories
        #ExSummary:Shows how to specify a set of categories for TOA fields.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # TOA fields can filter their entries by categories defined in this collection.
        toa_categories = aw.fields.ToaCategories()
        doc.field_options.toa_categories = toa_categories
        # This collection of categories comes with default values, which we can overwrite with custom values.
        self.assertEqual('Cases', toa_categories[1])
        self.assertEqual('Statutes', toa_categories[2])
        toa_categories[1] = 'My Category 1'
        toa_categories[2] = 'My Category 2'
        # We can always access the default values via this collection.
        self.assertEqual('Cases', aw.fields.ToaCategories.default_categories[1])
        self.assertEqual('Statutes', aw.fields.ToaCategories.default_categories[2])
        # Insert 2 TOA fields. TOA fields create an entry for each TA field in the document.
        # Use the "\c" switch to select the index of a category from our collection.
        #  With this switch, a TOA field will only pick up entries from TA fields that
        # also have a "\c" switch with a matching category index. Each TOA field will also display
        # the name of the category that its "\c" switch points to.
        builder.insert_field(field_code='TOA \\c 1 \\h', field_value=None)
        builder.insert_field(field_code='TOA \\c 2 \\h', field_value=None)
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        # Insert TOA entries across 2 categories. Our first TOA field will receive one entry,
        # from the second TA field whose "\c" switch also points to the first category.
        # The second TOA field will have two entries from the other two TA fields.
        builder.insert_field(field_code='TA \\c 2 \\l "entry 1"')
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.insert_field(field_code='TA \\c 1 \\l "entry 2"')
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.insert_field(field_code='TA \\c 2 \\l "entry 3"')
        doc.update_fields()
        doc.save(file_name=ARTIFACTS_DIR + 'FieldOptions.TOA.Categories.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'FieldOptions.TOA.Categories.docx')
        self.assertIsNone(doc.field_options.toa_categories)
        test_util.TestUtil.verify_field(expected_type=aw.fields.FieldType.FIELD_TOA, expected_field_code='TOA \\c 1 \\h', expected_result='My Category 1\rentry 2\t3\r', field=doc.range.fields[0])
        test_util.TestUtil.verify_field(expected_type=aw.fields.FieldType.FIELD_TOA, expected_field_code='TOA \\c 2 \\h', expected_result='My Category 2\r' + 'entry 1\t2\r' + 'entry 3\t4\r', field=doc.range.fields[1])