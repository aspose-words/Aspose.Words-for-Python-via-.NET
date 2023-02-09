# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

from typing import List

import aspose.words as aw

from api_example_base import ApiExampleBase, ARTIFACTS_DIR

class ExMailMergeCustomNested(ApiExampleBase):

    #ExStart
    #ExFor:MailMerge.execute_with_regions(IMailMergeDataSource)
    #ExSummary:Shows how to use mail merge regions to execute a nested mail merge.
    def test_custom_data_source(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Normally, MERGEFIELDs contain the name of a column of a mail merge data source.
        # Instead, we can use "TableStart:" and "TableEnd:" prefixes to begin/end a mail merge region.
        # Each region will belong to a table with a name that matches the string immediately after the prefix's colon.
        builder.insert_field(" MERGEFIELD TableStart:Customers")

        # These MERGEFIELDs are inside the mail merge region of the "Customers" table.
        # When we execute the mail merge, this field will receive data from rows in a data source named "Customers".
        builder.write("Full name:\t")
        builder.insert_field(" MERGEFIELD FullName ")
        builder.write("\nAddress:\t")
        builder.insert_field(" MERGEFIELD Address ")
        builder.write("\nOrders:\n")

        # Create a second mail merge region inside the outer region for a data source named "Orders".
        # The "Orders" data entries have a many-to-one relationship with the "Customers" data source.
        builder.insert_field(" MERGEFIELD TableStart:Orders")

        builder.write("\tItem name:\t")
        builder.insert_field(" MERGEFIELD Name ")
        builder.write("\n\tQuantity:\t")
        builder.insert_field(" MERGEFIELD Quantity ")
        builder.insert_paragraph()

        builder.insert_field(" MERGEFIELD TableEnd:Orders")
        builder.insert_field(" MERGEFIELD TableEnd:Customers")

        # Create related data with names that match those of our mail merge regions.
        customers = []
        customers.append(ExMailMergeCustomNested.Customer("Thomas Hardy", "120 Hanover Sq., London"))
        customers.append(ExMailMergeCustomNested.Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"))

        customers[0].orders.append(ExMailMergeCustomNested.Order("Rugby World Cup Cap", 2))
        customers[0].orders.append(ExMailMergeCustomNested.Order("Rugby World Cup Ball", 1))
        customers[1].orders.append(ExMailMergeCustomNested.Order("Rugby World Cup Guide", 1))

        # To mail merge from your data source, we must wrap it into an object that implements the IMailMergeDataSource interface.
        customers_data_source = ExMailMergeCustomNested.CustomerMailMergeDataSource(customers)

        doc.mail_merge.execute_with_regions(customers_data_source)

        doc.save(ARTIFACTS_DIR + "NestedMailMergeCustom.custom_data_source.docx")
        self._test_custom_data_source(customers, aw.Document(ARTIFACTS_DIR + "NestedMailMergeCustom.custom_data_source.docx")) #ExSkip

    class Customer:
        """An example of a "data entity" class in your application."""

        def __init__(self, full_name: str, address: str):

            self.full_name = full_name
            self.address = address
            self.orders = [] # type: List[Order]

    class Order:
        """An example of a child "data entity" class in your application."""

        def __init__(self, name: str, quantity: int):

            self.name = name
            self.quantity = quantity

    class CustomerMailMergeDataSource(aw.mailmerging.IMailMergeDataSource):
        """A custom mail merge data source that you implement to allow Aspose.Words
        to mail merge data from your Customer objects into Microsoft Word documents."""

        def __init__(self, customers):

            self.customers = customers

            # When we initialize the data source, its position must be before the first record.
            self.record_index = -1

        @property
        def table_name(self):
            """The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions."""

            return "Customers"

        def get_value(self, field_name: str):
            """Aspose.Words calls this method to get a value for every data field."""

            if field_name == "FullName":
                return self.customers[self.record_index].full_name

            if field_name == "Address":
                return self.customers[self.record_index].address

            if field_name == "Order":
                return self.customers[self.record_index].orders

            # Return "None" to the Aspose.Words mail merge engine to signify
            # that we could not find a field with this name.
            return None

        def move_next(self):
            """A standard implementation for moving to a next record in a collection."""

            if not self.is_eof:
                self.record_index += 1

            return not self.is_eof

        def get_child_data_source(self, table_name: str):

            # Get the child data source, whose name matches the mail merge region that uses its columns.
            if table_name == "Orders":
                return ExMailMergeCustomNested.OrderMailMergeDataSource(self.ustomers[self.record_index].orders)

            return None

        @property
        def is_eof(self) -> bool:

            return self.record_index >= len(self.customers)

    class OrderMailMergeDataSource(aw.mailmerging.IMailMergeDataSource):

        def __init__(self, orders):

            self.orders = orders

            # When we initialize the data source, its position must be before the first record.
            self.record_index = -1

        @property
        def table_name(self) -> str:
            """The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions."""

            return "Orders"

        def get_value(self, field_name: str):
            """Aspose.Words calls this method to get a value for every data field."""

            if field_name == "Name":
                return self.orders[self.record_index].name

            if field_name == "Quantity":
                return self.orders[self.record_index].quantity

            # Return "None" to the Aspose.Words mail merge engine to signify
            # that we could not find a field with this name.
            return None

        def move_next(self) -> bool:
            """A standard implementation for moving to a next record in a collection."""

            if not self.is_eof:
                self.record_index += 1

            return not self.is_eof

        def get_child_data_source(self, table_name: str) -> aw.mailmerging.IMailMergeDataSource:
            """Return None because we do not have any child elements for this sort of object."""

            return None

        @property
        def is_eof(self) -> bool:

            return self.record_index >= len(self.orders)

    #ExEnd

    def _test_custom_data_source(self, customers, doc: aw.Document):

        mail_merge_data: List[List[str]] = []

        for customer in customers:
            for order in customer.orders:
                mail_merge_data.add([order.name, order.quantity.to_string()])
            mail_merge_data.add([customer.full_name, customer.address])

        self.mail_merge_matches_array(mail_merge_data.to_array(), doc, False)
