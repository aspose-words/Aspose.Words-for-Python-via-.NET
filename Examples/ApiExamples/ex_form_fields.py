# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
from document_helper import DocumentHelper
import aspose.pydrawing
import aspose.words as aw
import aspose.words.fields
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExFormFields(ApiExampleBase):

    def test_create(self):
        #ExStart
        #ExFor:FormField
        #ExFor:FormField.result
        #ExFor:FormField.type
        #ExFor:FormField.name
        #ExSummary:Shows how to insert a combo box.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.write('Please select a fruit: ')
        # Insert a combo box which will allow a user to choose an option from a collection of strings.
        combo_box = builder.insert_combo_box('MyComboBox', ['Apple', 'Banana', 'Cherry'], 0)
        self.assertEqual('MyComboBox', combo_box.name)
        self.assertEqual(aw.fields.FieldType.FIELD_FORM_DROP_DOWN, combo_box.type)
        self.assertEqual('Apple', combo_box.result)
        # The form field will appear in the form of a "select" html tag.
        doc.save(file_name=ARTIFACTS_DIR + 'FormFields.Create.html')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'FormFields.Create.html')
        combo_box = doc.range.form_fields[0]
        self.assertEqual('MyComboBox', combo_box.name)
        self.assertEqual(aw.fields.FieldType.FIELD_FORM_DROP_DOWN, combo_box.type)
        self.assertEqual('Apple', combo_box.result)

    def test_text_input(self):
        #ExStart
        #ExFor:DocumentBuilder.insert_text_input
        #ExSummary:Shows how to insert a text input form field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.write('Please enter text here: ')
        # Insert a text input field, which will allow the user to click it and enter text.
        # Assign some placeholder text that the user may overwrite and pass
        # a maximum text length of 0 to apply no limit on the form field's contents.
        builder.insert_text_input('TextInput1', aw.fields.TextFormFieldType.REGULAR, '', 'Placeholder text', 0)
        # The form field will appear in the form of an "input" html tag, with a type of "text".
        doc.save(file_name=ARTIFACTS_DIR + 'FormFields.TextInput.html')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'FormFields.TextInput.html')
        text_input = doc.range.form_fields[0]
        self.assertEqual('TextInput1', text_input.name)
        self.assertEqual(aw.fields.TextFormFieldType.REGULAR, text_input.text_input_type)
        self.assertEqual('', text_input.text_input_format)
        self.assertEqual('Placeholder text', text_input.result)
        self.assertEqual(0, text_input.max_length)

    def test_delete_form_field(self):
        #ExStart
        #ExFor:FormField.remove_field
        #ExSummary:Shows how to delete a form field.
        doc = aw.Document(file_name=MY_DIR + 'Form fields.docx')
        form_field = doc.range.form_fields[3]
        form_field.remove_field()
        #ExEnd
        form_field_after = doc.range.form_fields[3]
        self.assertIsNone(form_field_after)

    def test_delete_form_field_associated_with_bookmark(self):
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.start_bookmark('MyBookmark')
        builder.insert_text_input('TextInput1', aw.fields.TextFormFieldType.REGULAR, '_test_form_field', 'SomeText', 0)
        builder.end_bookmark('MyBookmark')
        doc = DocumentHelper.save_open(doc)
        bookmark_before_delete_form_field = doc.range.bookmarks
        self.assertEqual('MyBookmark', bookmark_before_delete_form_field[0].name)
        form_field = doc.range.form_fields[0]
        form_field.remove_field()
        bookmark_after_delete_form_field = doc.range.bookmarks
        self.assertEqual('MyBookmark', bookmark_after_delete_form_field[0].name)

    def test_form_field_font_formatting(self):
        #ExStart
        #ExFor:FormField
        #ExSummary:Shows how to formatting the entire FormField, including the field value.
        doc = aw.Document(MY_DIR + 'Form fields.docx')
        form_field = doc.range.form_fields[0]
        form_field.font.bold = True
        form_field.font.size = 24
        form_field.font.color = aspose.pydrawing.Color.red
        form_field.result = 'Aspose.FormField'
        doc = DocumentHelper.save_open(doc)
        form_field_run = doc.first_section.body.first_paragraph.runs[1]
        self.assertEqual('Aspose.FormField', form_field_run.text)
        self.assertEqual(True, form_field_run.font.bold)
        self.assertEqual(24, form_field_run.font.size)
        self.assertEqual(aspose.pydrawing.Color.red.to_argb(), form_field_run.font.color.to_argb())
        #ExEnd

    def test_drop_down_item_collection(self):
        #ExStart
        #ExFor:DropDownItemCollection
        #ExFor:DropDownItemCollection.add(str)
        #ExFor:DropDownItemCollection.clear
        #ExFor:DropDownItemCollection.contains(str)
        #ExFor:DropDownItemCollection.count
        #ExFor:DropDownItemCollection.__iter__
        #ExFor:DropDownItemCollection.index_of(str)
        #ExFor:DropDownItemCollection.insert(int,str)
        #ExFor:DropDownItemCollection.__getitem__(int)
        #ExFor:DropDownItemCollection.remove(str)
        #ExFor:DropDownItemCollection.remove_at(int)
        #ExSummary:Shows how to insert a combo box field, and edit the elements in its item collection.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        # Insert a combo box, and then verify its collection of drop-down items.
        # In Microsoft Word, the user will click the combo box,
        # and then choose one of the items of text in the collection to display.
        items = ['One', 'Two', 'Three']
        combo_box_field = builder.insert_combo_box('DropDown', items, 0)
        drop_down_items = combo_box_field.drop_down_items
        self.assertEqual(3, drop_down_items.count)
        self.assertEqual('One', drop_down_items[0])
        self.assertEqual(1, drop_down_items.index_of('Two'))
        self.assertTrue(drop_down_items.contains('Three'))
        # There are two ways of adding a new item to an existing collection of drop-down box items.
        # 1 -  Append an item to the end of the collection:
        drop_down_items.add('Four')
        # 2 -  Insert an item before another item at a specified index:
        drop_down_items.insert(3, 'Three and a half')
        self.assertEqual(5, drop_down_items.count)
        # Iterate over the collection and print every element.
        for drop_down in drop_down_items:
            print(drop_down)
        # There are two ways of removing elements from a collection of drop-down items.
        # 1 -  Remove an item with contents equal to the passed string:
        drop_down_items.remove('Four')
        # 2 -  Remove an item at an index:
        drop_down_items.remove_at(3)
        self.assertEqual(3, drop_down_items.count)
        self.assertFalse(drop_down_items.contains('Three and a half'))
        self.assertFalse(drop_down_items.contains('Four'))
        doc.save(ARTIFACTS_DIR + 'FormFields.drop_down_item_collection.html')
        # Empty the whole collection of drop-down items.
        drop_down_items.clear()
        #ExEnd
        doc = DocumentHelper.save_open(doc)
        drop_down_items = doc.range.form_fields[0].drop_down_items
        self.assertEqual(0, drop_down_items.count)
        doc = aw.Document(ARTIFACTS_DIR + 'FormFields.drop_down_item_collection.html')
        drop_down_items = doc.range.form_fields[0].drop_down_items
        self.assertEqual(3, drop_down_items.count)
        self.assertEqual('One', drop_down_items[0])
        self.assertEqual('Two', drop_down_items[1])
        self.assertEqual('Three', drop_down_items[2])

    def _test_form_field(self, doc: aw.Document):
        doc = DocumentHelper.save_open(doc)
        fields = doc.range.fields
        self.assertEqual(3, fields.count)
        self.verify_field(aw.fields.FieldType.FIELD_FORM_DROP_DOWN, ' FORMDROPDOWN \x01', '', doc.range.fields[0])
        self.verify_field(aw.fields.FieldType.FIELD_FORM_CHECK_BOX, ' FORMCHECKBOX \x01', '', doc.range.fields[1])
        self.verify_field(aw.fields.FieldType.FIELD_FORM_TEXT_INPUT, ' FORMTEXT \x01', 'New placeholder text', doc.range.fields[2])
        form_fields = doc.range.form_fields
        self.assertEqual(3, form_fields.count)
        self.assertEqual(aw.fields.FieldType.FIELD_FORM_DROP_DOWN, form_fields[0].type)
        self.assertEqual(['One', 'Two', 'Three'], form_fields[0].drop_down_items)
        self.assertTrue(form_fields[0].calculate_on_exit)
        self.assertEqual(0, form_fields[0].drop_down_selected_index)
        self.assertTrue(form_fields[0].enabled)
        self.assertEqual('One', form_fields[0].result)
        self.assertEqual(aw.fields.FieldType.FIELD_FORM_CHECK_BOX, form_fields[1].type)
        self.assertTrue(form_fields[1].is_check_box_exact_size)
        self.assertEqual('Right click to check this box', form_fields[1].help_text)
        self.assertTrue(form_fields[1].own_help)
        self.assertEqual('Checkbox status text', form_fields[1].status_text)
        self.assertTrue(form_fields[1].own_status)
        self.assertEqual(50.0, form_fields[1].check_box_size)
        self.assertFalse(form_fields[1].checked)
        self.assertFalse(form_fields[1].default)
        self.assertEqual('0', form_fields[1].result)
        self.assertEqual(aw.fields.FieldType.FIELD_FORM_TEXT_INPUT, form_fields[2].type)
        self.assertEqual('EntryMacro', form_fields[2].entry_macro)
        self.assertEqual('ExitMacro', form_fields[2].exit_macro)
        self.assertEqual('Regular', form_fields[2].text_input_default)
        self.assertEqual('FIRST CAPITAL', form_fields[2].text_input_format)
        self.assertEqual(aw.fields.TextFormFieldType.REGULAR, form_fields[2].text_input_type)
        self.assertEqual(50, form_fields[2].max_length)
        self.assertEqual('New placeholder text', form_fields[2].result)