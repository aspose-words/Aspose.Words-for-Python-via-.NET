# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import aspose.words as aw

from api_example_base import ApiExampleBase

class ExVariableCollection(ApiExampleBase):

    def test_primer(self):

        #ExStart
        #ExFor:Document.variables
        #ExFor:VariableCollection
        #ExFor:VariableCollection.add
        #ExFor:VariableCollection.clear
        #ExFor:VariableCollection.contains
        #ExFor:VariableCollection.count
        #ExFor:VariableCollection.__iter__
        #ExFor:VariableCollection.index_of_key
        #ExFor:VariableCollection.remove
        #ExFor:VariableCollection.remove_at
        #ExSummary:Shows how to work with a document's variable collection.
        doc = aw.Document()
        variables = doc.variables

        # Every document has a collection of key/value pair variables, which we can add items to.
        variables.add("Home address", "123 Main St.")
        variables.add("City", "London")
        variables.add("Bedrooms", "3")

        self.assertEqual(3, variables.count)

        # We can display the values of variables in the document body using DOCVARIABLE fields.
        builder = aw.DocumentBuilder(doc)
        field = builder.insert_field(aw.fields.FieldType.FIELD_DOC_VARIABLE, True)
        field = field.as_field_doc_variable()
        field.variable_name = "Home address"
        field.update()

        self.assertEqual("123 Main St.", field.result)

        # Assigning values to existing keys will update them.
        variables.add("Home address", "456 Queen St.")

        # We will then have to update DOCVARIABLE fields to ensure they display an up-to-date value.
        self.assertEqual("123 Main St.", field.result)

        field.update()

        self.assertEqual("456 Queen St.", field.result)

        # Verify that the document variables with a certain name or value exist.
        self.assertTrue(variables.contains("City"))
        self.assertTrue(any(var.value == "London" for var in variables))

        # The collection of variables automatically sorts variables alphabetically by name.
        self.assertEqual(0, variables.index_of_key("Bedrooms"))
        self.assertEqual(1, variables.index_of_key("City"))
        self.assertEqual(2, variables.index_of_key("Home address"))

        # Enumerate over the collection of variables.
        for entry in doc.variables:
            print(f"Name: {entry.key}, Value: {entry.value}")

        # Below are three ways of removing document variables from a collection.
        # 1 -  By name:
        variables.remove("City")

        self.assertFalse(variables.contains("City"))

        # 2 -  By index:
        variables.remove_at(1)

        self.assertFalse(variables.contains("Home address"))

        # 3 -  Clear the whole collection at once:
        variables.clear()

        self.assertEqual(0, variables.count)
        #ExEnd
