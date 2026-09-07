import aspose.words as aw
from docs_examples_base import DocsExamplesBase, ARTIFACTS_DIR

class NestedMailMergeCustom(DocsExamplesBase):

    def test_custom_mail_merge(self):

        #ExStart:NestedMailMergeCustom
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.insert_field(" MERGEFIELD TableStart:Customer")

        builder.write("Full name:\t")
        builder.insert_field(" MERGEFIELD FullName ")
        builder.write("\nAddress:\t")
        builder.insert_field(" MERGEFIELD Address ")
        builder.write("\nOrders:\n")

        builder.insert_field(" MERGEFIELD TableStart:Order")

        builder.write("\tItem name:\t")
        builder.insert_field(" MERGEFIELD Name ")
        builder.write("\n\tQuantity:\t")
        builder.insert_field(" MERGEFIELD Quantity ")
        builder.insert_paragraph()

        builder.insert_field(" MERGEFIELD TableEnd:Order")

        builder.insert_field(" MERGEFIELD TableEnd:Customer")

        customers = [
            Customer("Thomas Hardy", "120 Hanover Sq., London"),
            Customer("Paolo Accorti", "Via Monte Bianco 34, Torino")
        ]

        customers[0].orders.append(Order("Rugby World Cup Cap", 2))
        customers[0].orders.append(Order("Rugby World Cup Ball", 1))
        customers[1].orders.append(Order("Rugby World Cup Guide", 1))

        # To be able to mail merge from your data source,
        # it must be wrapped into an object that implements the IMailMergeDataSource interface.
        customers_data_source = CustomerMailMergeDataSource(customers)

        doc.mail_merge.execute_with_regions(customers_data_source)

        doc.save(ARTIFACTS_DIR + "NestedMailMergeCustom.custom_mail_merge.docx")
        #ExEnd:NestedMailMergeCustom

class Customer:
    """An example of a "data entity" class in your application."""

    def __init__(self, full_name: str, address: str):
        self.full_name = full_name
        self.address = address
        self.orders = []

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

        # When the data source is initialized, it must be positioned before the first record.
        self.record_index = -1

    @property
    def table_name(self):
        """The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions."""
        return "Customer"

    def get_value(self, field_name: str, field_value):
        """Aspose.Words calls this method to get a value for every data field.
        The value is returned through the field_value list."""
        if field_name == "FullName":
            field_value[0] = self.customers[self.record_index].full_name
            return True
        if field_name == "Address":
            field_value[0] = self.customers[self.record_index].address
            return True

        return False

    def move_next(self):
        """A standard implementation for moving to a next record in a collection."""
        if not self.is_eof:
            self.record_index += 1

        return not self.is_eof

    #ExStart:GetChildDataSource
    #GistId:41abb81bce75f4ca71861895e5fc69c4
    def get_child_data_source(self, table_name: str):
        # Get the child collection to merge it with the region provided with the table_name variable.
        if table_name == "Order":
            return OrderMailMergeDataSource(self.customers[self.record_index].orders)

        return None
    #ExEnd:GetChildDataSource

    @property
    def is_eof(self):
        return self.record_index >= len(self.customers)

class OrderMailMergeDataSource(aw.mailmerging.IMailMergeDataSource):

    def __init__(self, orders):
        self.orders = orders

        # When the data source is initialized, it must be positioned before the first record.
        self.record_index = -1

    @property
    def table_name(self):
        """The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions."""
        return "Order"

    def get_value(self, field_name: str, field_value):
        """Aspose.Words calls this method to get a value for every data field."""
        if field_name == "Name":
            field_value[0] = self.orders[self.record_index].name
            return True
        if field_name == "Quantity":
            field_value[0] = self.orders[self.record_index].quantity
            return True

        return False

    def move_next(self):
        """A standard implementation for moving to a next record in a collection."""
        if not self.is_eof:
            self.record_index += 1

        return not self.is_eof

    def get_child_data_source(self, table_name: str):
        # Return None because we haven't any child elements for this sort of object.
        return None

    @property
    def is_eof(self):
        return self.record_index >= len(self.orders)
