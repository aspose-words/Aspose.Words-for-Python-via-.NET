import xml.etree.ElementTree as ET

import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class WorkingWithXmlData(DocsExamplesBase):

    def test_xml_to_mail_merge(self):

        order_xml = ET.parse(MY_DIR + "Mail merge data - Purchase order.xml").getroot()

        # Query the purchase order XML file to extract the order items into objects of a simple type.
        #
        # Ensure you give the type attributes the same names as the MERGEFIELD fields in the document.
        #
        # To pass the actual values stored in the XML element or attribute to Aspose.Words,
        # we read them as strings. This prevents the XML tags from being inserted into the final document.

        #ExStart:LinqToXmlMailMergeOrderItems
        order_items = [
            OrderItem(
                part_number=item.get("PartNumber"),
                product_name=self.element_text(item, "ProductName"),
                quantity=self.element_text(item, "Quantity"),
                us_price=self.element_text(item, "USPrice"),
                comment=self.element_text(item, "Comment"),
                ship_date=self.element_text(item, "ShipDate"))
            for item in order_xml.iter("Item")]
        #ExEnd:LinqToXmlMailMergeOrderItems

        #ExStart:LinqToXmlQueryForDeliveryAddress
        delivery_address = [
            DeliveryAddress(
                name=self.element_text(address, "Name"),
                country=self.element_text(address, "Country"),
                zip=self.element_text(address, "Zip"),
                state=self.element_text(address, "State"),
                city=self.element_text(address, "City"),
                street=self.element_text(address, "Street"))
            for address in order_xml.findall("Address")
            if address.get("Type") == "Shipping"]
        #ExEnd:LinqToXmlQueryForDeliveryAddress

        order_items_data_source = MyMailMergeDataSource(order_items, "Items")
        delivery_data_source = MyMailMergeDataSource(delivery_address)

        #ExStart:LinqToXmlMailMerge
        doc = aw.Document(MY_DIR + "Mail merge destinations - LINQ.docx")

        # Fill the document with data from our data sources using mail merge regions for populating the order items
        # table is required because it allows the region to be repeated in the document for each order item.
        doc.mail_merge.execute_with_regions(order_items_data_source)

        doc.mail_merge.execute(delivery_data_source)

        doc.save(ARTIFACTS_DIR + "WorkingWithXmlData.xml_mail_merge.docx")
        #ExEnd:LinqToXmlMailMerge

    def test_mustache_syntax_using_custom_data_source(self):

        #ExStart:MailMergeUsingMustacheSyntax
        doc = aw.Document(MY_DIR + "Mail merge destinations - Vendor.docx")

        # Fill the data source with the records to repeat inside the "{{#foreach Vendor}}" section.
        vendors = [ListItem(f"Vendor {i}") for i in range(1, 4)]

        # Activate performing a mail merge operation into additional field types.
        doc.mail_merge.use_non_merge_fields = True

        doc.mail_merge.execute_with_regions(MyMailMergeDataSource(vendors, "Vendor"))

        doc.save(ARTIFACTS_DIR + "WorkingWithXmlData.mustache_syntax_using_custom_data_source.docx")
        #ExEnd:MailMergeUsingMustacheSyntax

    @staticmethod
    def element_text(parent, name):
        """Returns the text of a child element, or an empty string if the element is missing."""
        element = parent.find(name)
        return element.text if element is not None else ""

class OrderItem:
    """A "data entity" whose attribute names match the merge field names in the template."""

    def __init__(self, part_number, product_name, quantity, us_price, comment, ship_date):
        self.PartNumber = part_number
        self.ProductName = product_name
        self.Quantity = quantity
        self.USPrice = us_price
        self.Comment = comment
        self.ShipDate = ship_date

class ListItem:
    """A "data entity" used to populate a mustache section."""

    def __init__(self, name):
        self.Name = name

class DeliveryAddress:
    """A "data entity" whose attribute names match the merge field names in the template."""

    def __init__(self, name, country, zip, state, city, street):
        self.Name = name
        self.Country = country
        self.Zip = zip
        self.State = state
        self.City = city
        self.Street = street

#ExStart:MyMailMergeDataSource
class MyMailMergeDataSource(aw.mailmerging.IMailMergeDataSource):
#ExEnd:MyMailMergeDataSource
    """Aspose.Words does not accept collections of arbitrary objects as input for mail merge directly,
    but provides a generic mechanism that allows mail merges from any data source.

    This class is a simple implementation of the Aspose.Words custom mail merge data source
    interface that accepts any iterable object.
    Aspose.Words calls this class during the mail merge to retrieve the data."""

    #ExStart:MyMailMergeDataSourceConstructor
    def __init__(self, data, table_name=""):
        """Creates a new instance of a custom mail merge data source.

        :param data: Any iterable collection of data entities.
        :param table_name: The name of the data source is only used when you perform a mail merge
            with regions. If you prefer to use the simple mail merge, then omit this parameter."""
        self.records = list(data)
        self._table_name = table_name

        # When the data source is initialized, it must be positioned before the first record.
        self.record_index = -1
    #ExEnd:MyMailMergeDataSourceConstructor

    #ExStart:MyMailMergeDataSourceTableName
    @property
    def table_name(self):
        """The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions."""
        return self._table_name
    #ExEnd:MyMailMergeDataSourceTableName

    #ExStart:MyMailMergeDataSourceGetValue
    def get_value(self, field_name: str, field_value):
        """Aspose.Words calls this method to get a value for every data field.

        This is a simple "generic" implementation of a data source that can work over any collection.
        This implementation assumes that the merge field name in the document matches the attribute's name
        on the object in the collection and uses getattr to get the attribute's value."""
        record = self.records[self.record_index]

        if hasattr(record, field_name):
            # Aspose.Words returns the value through the field_value list.
            field_value[0] = getattr(record, field_name)
            return True

        return False
    #ExEnd:MyMailMergeDataSourceGetValue

    #ExStart:MyMailMergeDataSourceMoveNext
    def move_next(self):
        """Moves to the next record in the collection."""
        if not self.is_eof:
            self.record_index += 1

        return not self.is_eof
    #ExEnd:MyMailMergeDataSourceMoveNext

    def get_child_data_source(self, table_name: str):
        return None

    @property
    def is_eof(self):
        return self.record_index >= len(self.records)
