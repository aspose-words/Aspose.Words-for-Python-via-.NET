# Copyright (c) 2001-2022 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR
from document_helper import DocumentHelper

class ExFormFields(ApiExampleBase):

    def test_create(self):

        #ExStart
        #ExFor:FormField
        #ExFor:FormField.result
        #ExFor:FormField.type
        #ExFor:FormField.name
        #ExSummary:Shows how to insert a combo box.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Please select a fruit: ")

        # Insert a combo box which will allow a user to choose an option from a collection of strings.
        combo_box = builder.insert_combo_box("MyComboBox", ["Apple", "Banana", "Cherry"], 0)

        self.assertEqual("MyComboBox", combo_box.name)
        self.assertEqual(aw.fields.FieldType.FIELD_FORM_DROP_DOWN, combo_box.type)
        self.assertEqual("Apple", combo_box.result)

        # The form field will appear in the form of a "select" html tag.
        doc.save(ARTIFACTS_DIR + "FormFields.create.html")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "FormFields.create.html")
        combo_box = doc.range.form_fields[0]

        self.assertEqual("MyComboBox", combo_box.name)
        self.assertEqual(aw.fields.FieldType.FIELD_FORM_DROP_DOWN, combo_box.type)
        self.assertEqual("Apple", combo_box.result)

    def test_text_input(self):

        #ExStart
        #ExFor:DocumentBuilder.insert_text_input
        #ExSummary:Shows how to insert a text input form field.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Please enter text here: ")

        # Insert a text input field, which will allow the user to click it and enter text.
        # Assign some placeholder text that the user may overwrite and pass
        # a maximum text length of 0 to apply no limit on the form field's contents.
        builder.insert_text_input("TextInput1", aw.fields.TextFormFieldType.REGULAR, "", "Placeholder text", 0)

        # The form field will appear in the form of an "input" html tag, with a type of "text".
        doc.save(ARTIFACTS_DIR + "FormFields.text_input.html")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "FormFields.text_input.html")

        text_input = doc.range.form_fields[0]

        self.assertEqual("TextInput1", text_input.name)
        self.assertEqual(aw.fields.TextFormFieldType.REGULAR, text_input.text_input_type)
        self.assertEqual("", text_input.text_input_format)
        self.assertEqual("Placeholder text", text_input.result)
        self.assertEqual(0, text_input.max_length)

    def test_delete_form_field(self):

        #ExStart
        #ExFor:FormField.remove_field
        #ExSummary:Shows how to delete a form field.
        doc = aw.Document(MY_DIR + "Form fields.docx")

        form_field = doc.range.form_fields[3]
        form_field.remove_field()
        #ExEnd

        form_field_after = doc.range.form_fields[3]

        self.assertIsNone(form_field_after)

    def test_delete_form_field_associated_with_bookmark(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_bookmark("MyBookmark")
        builder.insert_text_input("TextInput1", aw.fields.TextFormFieldType.REGULAR, "_test_form_field", "SomeText", 0)
        builder.end_bookmark("MyBookmark")

        doc = DocumentHelper.save_open(doc)

        bookmark_before_delete_form_field = doc.range.bookmarks
        self.assertEqual("MyBookmark", bookmark_before_delete_form_field[0].name)

        form_field = doc.range.form_fields[0]
        form_field.remove_field()

        bookmark_after_delete_form_field = doc.range.bookmarks
        self.assertEqual("MyBookmark", bookmark_after_delete_form_field[0].name)

    def test_form_field_font_formatting(self):

        #ExStart
        #ExFor:FormField
        #ExSummary:Shows how to formatting the entire FormField, including the field value.
        doc = aw.Document(MY_DIR + "Form fields.docx")

        form_field = doc.range.form_fields[0]
        form_field.font.bold = True
        form_field.font.size = 24
        form_field.font.color = drawing.Color.red

        form_field.result = "Aspose.FormField"

        doc = DocumentHelper.save_open(doc)

        form_field_run = doc.first_section.body.first_paragraph.runs[1]

        self.assertEqual("Aspose.FormField", form_field_run.text)
        self.assertEqual(True, form_field_run.font.bold)
        self.assertEqual(24, form_field_run.font.size)
        self.assertEqual(drawing.Color.red.to_argb(), form_field_run.font.color.to_argb())
        #ExEnd

    ##ExStart
    ##ExFor:FormField.accept(DocumentVisitor)
    ##ExFor:FormField.calculate_on_exit
    ##ExFor:FormField.check_box_size
    ##ExFor:FormField.checked
    ##ExFor:FormField.default
    ##ExFor:FormField.drop_down_items
    ##ExFor:FormField.drop_down_selected_index
    ##ExFor:FormField.enabled
    ##ExFor:FormField.entry_macro
    ##ExFor:FormField.exit_macro
    ##ExFor:FormField.help_text
    ##ExFor:FormField.is_check_box_exact_size
    ##ExFor:FormField.max_length
    ##ExFor:FormField.own_help
    ##ExFor:FormField.own_status
    ##ExFor:FormField.set_text_input_value(object)
    ##ExFor:FormField.status_text
    ##ExFor:FormField.text_input_default
    ##ExFor:FormField.text_input_format
    ##ExFor:FormField.text_input_type
    ##ExFor:FormFieldCollection
    ##ExFor:FormFieldCollection.clear
    ##ExFor:FormFieldCollection.count
    ##ExFor:FormFieldCollection.__iter__
    ##ExFor:FormFieldCollection.__getitem__(int)
    ##ExFor:FormFieldCollection.__getitem__(str)
    ##ExFor:FormFieldCollection.remove(str)
    ##ExFor:FormFieldCollection.remove_at(int)
    ##ExFor:Range.form_fields
    ##ExSummary:Shows how insert different kinds of form fields into a document, and process them with using a document visitor implementation.
    #def test_visitor(self):

    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    # Use a document builder to insert a combo box.
    #    builder.write("Choose a value from this combo box: ")
    #    combo_box = builder.insert_combo_box("MyComboBox", ["One", "Two", "Three"], 0)
    #    combo_box.calculate_on_exit = True
    #    self.assertEqual(3, combo_box.drop_down_items.count)
    #    self.assertEqual(0, combo_box.drop_down_selected_index)
    #    self.assertTrue(combo_box.enabled)

    #    builder.insert_break(aw.BreakType.PARAGRAPH_BREAK)

    #    # Use a document builder to insert a check box.
    #    builder.write("Click this check box to tick/untick it: ")
    #    check_box = builder.insert_check_box("MyCheckBox", False, 50)
    #    check_box.is_check_box_exact_size = True
    #    check_box.help_text = "Right click to check this box"
    #    check_box.own_help = True
    #    check_box.status_text = "Checkbox status text"
    #    check_box.own_status = True
    #    self.assertEqual(50.0, check_box.check_box_size)
    #    self.assertFalse(check_box.checked)
    #    self.assertFalse(check_box.default)

    #    builder.insert_break(aw.BreakType.PARAGRAPH_BREAK)

    #    # Use a document builder to insert text input form field.
    #    builder.write("Enter text here: ")
    #    text_input = builder.insert_text_input("MyTextInput", aw.fields.TextFormFieldType.REGULAR, "", "Placeholder text", 50)
    #    text_input.entry_macro = "EntryMacro"
    #    text_input.exit_macro = "ExitMacro"
    #    text_input.text_input_default = "Regular"
    #    text_input.text_input_format = "FIRST CAPITAL"
    #    text_input.set_text_input_value("New placeholder text")
    #    self.assertEqual(aw.fields.TextFormFieldType.REGULAR, text_input.text_input_type)
    #    self.assertEqual(50, text_input.max_length)

    #    # This collection contains all our form fields.
    #    form_fields = doc.range.form_fields
    #    self.assertEqual(3, form_fields.count)

    #    # Fields display our form fields. We can see their field codes by opening this document
    #    # in Microsoft and pressing Alt + F9. These fields have no switches,
    #    # and members of the FormField object fully govern their form fields' content.
    #    self.assertEqual(3, doc.range.fields.count)
    #    self.assertEqual(" FORMDROPDOWN \u0001", doc.range.fields[0].get_field_code())
    #    self.assertEqual(" FORMCHECKBOX \u0001", doc.range.fields[1].get_field_code())
    #    self.assertEqual(" FORMTEXT \u0001", doc.range.fields[2].get_field_code())

    #    # Allow each form field to accept a document visitor.
    #    form_field_visitor = ExFormFields.FormFieldVisitor()

    #    for field in form_fields:
    #        field.accept(form_field_visitor)

    #    print(form_field_visitor.get_text())

    #    doc.update_fields()
    #    doc.save(ARTIFACTS_DIR + "FormFields.visitor.html")
    #    self._test_form_field(doc) #ExSkip

    #class FormFieldVisitor(aw.DocumentVisitor):
    #    """Visitor implementation that prints details of form fields that it visits."""

    #    def __init__(self):
    #        self.builder = io.StringIO()

    #    def visit_form_field(self, form_field: aw.fields.FormField) -> aw.VisitorAction:
    #        """Called when a FormField node is encountered in the document."""

    #        self.append_line(form_field.type + ": \"" + form_field.name + "\"")
    #        self.append_line("\tStatus: " + ("Enabled" if form_field.enabled else "Disabled"))
    #        self.append_line("\tHelp Text:  " + form_field.help_text)
    #        self.append_line("\tEntry macro name: " + form_field.entry_macro)
    #        self.append_line("\tExit macro name: " + form_field.exit_macro)

    #        if form_field.type == aw.fields.FieldType.FIELD_FORM_DROP_DOWN:
    #            self.append_line("\tDrop-down items count: " + form_field.drop_down_items.count + ", default selected item index: " + form_field.drop_down_selected_index)
    #            self.append_line("\tDrop-down items: " + ", ".join(form_field.drop_down_items.to_array()))

    #        elif form_field.type == aw.fields.FieldType.FIELD_FORM_CHECK_BOX:
    #            self.append_line("\tCheckbox size: " + form_field.check_box_size)
    #            self.append_line("\t" + "Checkbox is currently: " + ("checked, " if form_field.checked else "unchecked, ") + "by default: " + ("checked" if form_field.default else "unchecked"))

    #        elif form_field.type == aw.fields.FieldType.FIELD_FORM_TEXT_INPUT:
    #            self.append_line("\tInput format: " + form_field.text_input_format)
    #            self.append_line("\tCurrent contents: " + form_field.result)

    #        # Let the visitor continue visiting other nodes.
    #        return aw.VisitorAction.CONTINUE

    #    def append_line(self, text: str):
    #        """Adds newline char-terminated text to the current output."""

    #        self.builder.write(text + '\n')

    #    def get_text(self) -> str:
    #        """Gets the plain text of the document that was accumulated by the visitor."""

    #        return self.builder.getvalue()

    ##ExEnd

    def _test_form_field(self, doc: aw.Document):

        doc = DocumentHelper.save_open(doc)
        fields = doc.range.fields
        self.assertEqual(3, fields.count)

        self.verify_field(aw.fields.FieldType.FIELD_FORM_DROP_DOWN, " FORMDROPDOWN \u0001", "", doc.range.fields[0])
        self.verify_field(aw.fields.FieldType.FIELD_FORM_CHECK_BOX, " FORMCHECKBOX \u0001", "", doc.range.fields[1])
        self.verify_field(aw.fields.FieldType.FIELD_FORM_TEXT_INPUT, " FORMTEXT \u0001", "New placeholder text", doc.range.fields[2])

        form_fields = doc.range.form_fields
        self.assertEqual(3, form_fields.count)

        self.assertEqual(aw.fields.FieldType.FIELD_FORM_DROP_DOWN, form_fields[0].type)
        self.assertEqual(["One", "Two", "Three"], form_fields[0].drop_down_items)
        self.assertTrue(form_fields[0].calculate_on_exit)
        self.assertEqual(0, form_fields[0].drop_down_selected_index)
        self.assertTrue(form_fields[0].enabled)
        self.assertEqual("One", form_fields[0].result)

        self.assertEqual(aw.fields.FieldType.FIELD_FORM_CHECK_BOX, form_fields[1].type)
        self.assertTrue(form_fields[1].is_check_box_exact_size)
        self.assertEqual("Right click to check this box", form_fields[1].help_text)
        self.assertTrue(form_fields[1].own_help)
        self.assertEqual("Checkbox status text", form_fields[1].status_text)
        self.assertTrue(form_fields[1].own_status)
        self.assertEqual(50.0, form_fields[1].check_box_size)
        self.assertFalse(form_fields[1].checked)
        self.assertFalse(form_fields[1].default)
        self.assertEqual("0", form_fields[1].result)

        self.assertEqual(aw.fields.FieldType.FIELD_FORM_TEXT_INPUT, form_fields[2].type)
        self.assertEqual("EntryMacro", form_fields[2].entry_macro)
        self.assertEqual("ExitMacro", form_fields[2].exit_macro)
        self.assertEqual("Regular", form_fields[2].text_input_default)
        self.assertEqual("FIRST CAPITAL", form_fields[2].text_input_format)
        self.assertEqual(aw.fields.TextFormFieldType.REGULAR, form_fields[2].text_input_type)
        self.assertEqual(50, form_fields[2].max_length)
        self.assertEqual("New placeholder text", form_fields[2].result)

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
        items = ["One", "Two", "Three"]
        combo_box_field = builder.insert_combo_box("DropDown", items, 0)
        drop_down_items = combo_box_field.drop_down_items

        self.assertEqual(3, drop_down_items.count)
        self.assertEqual("One", drop_down_items[0])
        self.assertEqual(1, drop_down_items.index_of("Two"))
        self.assertTrue(drop_down_items.contains("Three"))

        # There are two ways of adding a new item to an existing collection of drop-down box items.
        # 1 -  Append an item to the end of the collection:
        drop_down_items.add("Four")

        # 2 -  Insert an item before another item at a specified index:
        drop_down_items.insert(3, "Three and a half")

        self.assertEqual(5, drop_down_items.count)

        # Iterate over the collection and print every element.
        for drop_down in drop_down_items:
            print(drop_down)

        # There are two ways of removing elements from a collection of drop-down items.
        # 1 -  Remove an item with contents equal to the passed string:
        drop_down_items.remove("Four")

        # 2 -  Remove an item at an index:
        drop_down_items.remove_at(3)

        self.assertEqual(3, drop_down_items.count)
        self.assertFalse(drop_down_items.contains("Three and a half"))
        self.assertFalse(drop_down_items.contains("Four"))

        doc.save(ARTIFACTS_DIR + "FormFields.drop_down_item_collection.html")

        # Empty the whole collection of drop-down items.
        drop_down_items.clear()
        #ExEnd

        doc = DocumentHelper.save_open(doc)
        drop_down_items = doc.range.form_fields[0].drop_down_items

        self.assertEqual(0, drop_down_items.count)

        doc = aw.Document(ARTIFACTS_DIR + "FormFields.drop_down_item_collection.html")
        drop_down_items = doc.range.form_fields[0].drop_down_items

        self.assertEqual(3, drop_down_items.count)
        self.assertEqual("One", drop_down_items[0])
        self.assertEqual("Two", drop_down_items[1])
        self.assertEqual("Three", drop_down_items[2])
