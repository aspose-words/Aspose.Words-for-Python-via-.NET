# Copyright (c) 2001-2022 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

from datetime import datetime

import aspose.words as aw

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, GOLDS_DIR, DATABASE_DIR
from document_helper import DocumentHelper

class ExMailMerge(ApiExampleBase):

    def test_execute_array(self):

        response = None

        #ExStart
        #ExFor:MailMerge.execute(List[str],List[object])
        #ExFor:ContentDisposition
        #ExFor:Document.save(HttpResponse,str,ContentDisposition,SaveOptions)
        #ExSummary:Shows how to perform a mail merge, and then save the document to the client browser.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_field(" MERGEFIELD FullName ")
        builder.insert_paragraph()
        builder.insert_field(" MERGEFIELD Company ")
        builder.insert_paragraph()
        builder.insert_field(" MERGEFIELD Address ")
        builder.insert_paragraph()
        builder.insert_field(" MERGEFIELD City ")

        doc.mail_merge.execute(["FullName", "Company", "Address", "City"],
            ["James Bond", "MI5 Headquarters", "Milbank", "London"])

        # Send the document to the client browser.
        with self.assertRaises(Exception):
            #Thrown because HttpResponse is null in the test.
            doc.save(response, "Artifacts/MailMerge.execute_array.docx", aw.ContentDisposition.INLINE, None)

        # We will need to close this response manually to ensure that we do not add any superfluous content to the document after saving.
        with self.assertRaises(Exception):
            response.end()
        #ExEnd

        doc = DocumentHelper.save_open(doc)

        self.mail_merge_matches_array([["James Bond", "MI5 Headquarters", "Milbank", "London"]], doc, True)

    def test_execute_data_reader(self):

        #ExStart
        #ExFor:MailMerge.execute(IDataReader)
        #ExSummary:Shows how to run a mail merge using data from a data reader.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Product:\t")
        builder.insert_field(" MERGEFIELD ProductName")
        builder.write("\nSupplier:\t")
        builder.insert_field(" MERGEFIELD CompanyName")
        builder.writeln()
        builder.insert_field(" MERGEFIELD QuantityPerUnit")
        builder.write(" for f")
        builder.insert_field(" MERGEFIELD UnitPrice")

        # Create a connection string that points to the "Northwind" database file
        # in our local file system, open a connection, and set up an SQL query.
        connection_string = r"Driver={Microsoft Access Driver (*.mdb)};Dbq=" + DATABASE_DIR + "Northwind.mdb"
        query = """
            SELECT Products.ProductName, Suppliers.CompanyName, Products.QuantityPerUnit, {fn ROUND(Products.UnitPrice,2)} as UnitPrice
            FROM Products
            INNER JOIN Suppliers
            ON Products.SupplierID = Suppliers.SupplierID"""

        connection = OdbcConnection()
        connection.connection_string = connection_string
        connection.open()

        # Create an SQL command that will source data for our mail merge.
        # The names of the table's columns that this SELECT statement will return
        # will need to correspond to the merge fields we placed above.
        command = connection.create_command()
        command.command_text = query

        # This will run the command and store the data in the reader.
        reader = command.execute_reader(CommandBehavior.CLOSE_CONNECTION)

        # Take the data from the reader and use it in the mail merge.
        doc.mail_merge.execute(reader)

        doc.save(ARTIFACTS_DIR + "MailMerge.execute_data_reader.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "MailMerge.execute_data_reader.docx")

        self.mail_merge_matches_query_result(DATABASE_DIR + "Northwind.mdb", query, doc, True)

    #ExStart
    #ExFor:MailMerge.execute_ado(object)
    #ExSummary:Shows how to run a mail merge with data from an ADO dataset.
    def test_execute_ado(self):

        doc = ExMailMerge.create_source_doc_ado_mail_merge()

        # To work with ADO DataSets, we will need to add a reference to the Microsoft ActiveX Data Objects library,
        # which is included in the .NET distribution and stored in "adodb.dll".
        connection = ADODB.Connection()

        # Create a connection string that points to the "Northwind" database file
        # in our local file system and open a connection.
        connection_string = r"Provider=Microsoft.jet.oledb.4.0;Data Source=" + DATABASE_DIR + "Northwind.mdb"
        connection.open(connection_string)

        # Populate our DataSet by running an SQL command on our database.
        # The names of the columns in the result table will need to correspond
        # to the values of the MERGEFIELDS that will accommodate our data.
        command = r"SELECT ProductName, QuantityPerUnit, UnitPrice FROM Products"

        recordset = ADODB.Recordset()
        recordset.open(command, connection)

        # Execute the mail merge and save the document.
        doc.mail_merge.execute_ado(recordset)
        doc.save(ARTIFACTS_DIR + "MailMerge.execute_ado.docx")
        self.mail_merge_matches_query_result(DATABASE_DIR + "Northwind.mdb", command, doc, True) #ExSkip

    @staticmethod
    def create_source_doc_ado_mail_merge() -> aw.Document:
        """Create a blank document and populate it with MERGEFIELDS that will accept data when a mail merge is executed."""

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Product:\t")
        builder.insert_field(" MERGEFIELD ProductName")
        builder.writeln()
        builder.insert_field(" MERGEFIELD QuantityPerUnit")
        builder.write(" for f")
        builder.insert_field(" MERGEFIELD UnitPrice")

        return doc

    #ExEnd

    #ExStart
    #ExFor:MailMerge.execute_with_regions_ado(object,str)
    #ExSummary:Shows how to run a mail merge with multiple regions, compiled with data from an ADO dataset.
    def test_execute_with_regions_ado(self):

        doc = ExMailMerge.create_source_doc_ado_mail_merge_with_regions()

        # To work with ADO DataSets, we will need to add a reference to the Microsoft ActiveX Data Objects library,
        # which is included in the .NET distribution and stored in "adodb.dll".
        connection = ADODB.Connection()

        # Create a connection string that points to the "Northwind" database file
        # in our local file system and open a connection.
        connection_string = r"Provider=Microsoft.jet.oledb.4.0;Data Source=" + DATABASE_DIR + "Northwind.mdb"
        connection.open(connection_string)

        # Populate our DataSet by running an SQL command on our database.
        # The names of the columns in the result table will need to correspond
        # to the values of the MERGEFIELDS that will accommodate our data.
        command = "SELECT FirstName, LastName, City FROM Employees"

        recordset = ADODB.Recordset()
        recordset.open(command, connection)

        # Run a mail merge on just the first region, filling its MERGEFIELDS with data from the record set.
        doc.mail_merge.execute_with_regions_ado(recordset, "MergeRegion1")

        # Close the record set and reopen it with data from another SQL query.
        command = "SELECT * FROM Customers"

        recordset.close()
        recordset.open(command, connection)

        # Run a second mail merge on the second region and save the document.
        doc.mail_merge.execute_with_regions_ado(recordset, "MergeRegion2")

        doc.save(ARTIFACTS_DIR + "MailMerge.execute_with_regions_ado.docx")
        self.mail_merge_matches_query_result_multiple(DATABASE_DIR + "Northwind.mdb", ["SELECT FirstName, LastName, City FROM Employees", "SELECT ContactName, Address, City FROM Customers"], aw.Document(ARTIFACTS_DIR + "MailMerge.execute_with_regions_ado.docx"), False) #ExSkip

    @staticmethod
    def create_source_doc_ado_mail_merge_with_regions() -> aw.Document:
        """Create a document with two mail merge regions."""

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("\tEmployees: ")
        builder.insert_field(" MERGEFIELD TableStart:MergeRegion1")
        builder.insert_field(" MERGEFIELD FirstName")
        builder.write(", ")
        builder.insert_field(" MERGEFIELD LastName")
        builder.write(", ")
        builder.insert_field(" MERGEFIELD City")
        builder.insert_field(" MERGEFIELD TableEnd:MergeRegion1")
        builder.insert_paragraph()

        builder.writeln("\tCustomers: ")
        builder.insert_field(" MERGEFIELD TableStart:MergeRegion2")
        builder.insert_field(" MERGEFIELD ContactName")
        builder.write(", ")
        builder.insert_field(" MERGEFIELD Address")
        builder.write(", ")
        builder.insert_field(" MERGEFIELD City")
        builder.insert_field(" MERGEFIELD TableEnd:MergeRegion2")

        return doc

    #ExEnd

    #ExStart
    #ExFor:Document
    #ExFor:MailMerge
    #ExFor:MailMerge.execute(DataTable)
    #ExFor:MailMerge.execute(DataRow)
    #ExFor:Document.mail_merge
    #ExSummary:Shows how to execute a mail merge with data from a DataTable.
    def test_execute_data_table(self):

        table = DataTable("Test")
        table.columns.add("CustomerName")
        table.columns.add("Address")
        table.rows.add(["Thomas Hardy", "120 Hanover Sq., London"])
        table.rows.add(["Paolo Accorti", "Via Monte Bianco 34, Torino"])

        # Below are two ways of using a DataTable as the data source for a mail merge.
        # 1 -  Use the entire table for the mail merge to create one output mail merge document for every row in the table:
        doc = ExMailMerge.create_source_doc_execute_data_table()

        doc.mail_merge.execute(table)

        doc.save(ARTIFACTS_DIR + "MailMerge.execute_data_table.whole_table.docx")

        # 2 -  Use one row of the table to create one output mail merge document:
        doc = ExMailMerge.create_source_doc_execute_data_table()

        doc.mail_merge.execute(table.rows[1])

        doc.save(ARTIFACTS_DIR + "MailMerge.execute_data_table.one_row.docx")
        self._test_ado_data_table(aw.Document(ARTIFACTS_DIR + "MailMerge.execute_data_table.whole_table.docx"), aw.Document(ARTIFACTS_DIR + "MailMerge.execute_data_table.one_row.docx"), table) #ExSkip

    @staticmethod
    def create_source_doc_execute_data_table() -> aw.Document:
        """Creates a mail merge source document."""

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_field(" MERGEFIELD CustomerName ")
        builder.insert_paragraph()
        builder.insert_field(" MERGEFIELD Address ")

        return doc

    #ExEnd

    def _test_ado_data_table(self, doc_whole_table: aw.Document, doc_one_row: aw.Document, table: DataTable):

        self.mail_merge_matches_data_table(table, doc_whole_table, True)

        row_as_table = DataTable()
        row_as_table.import_row(table.rows[1])

        self.mail_merge_matches_data_table(row_as_table, doc_one_row, True)

    def test_execute_data_view(self):

        #ExStart
        #ExFor:MailMerge.execute(DataView)
        #ExSummary:Shows how to edit mail merge data with a DataView.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.write("Congratulations ")
        builder.insert_field(" MERGEFIELD Name")
        builder.write(" for passing with a grade of ")
        builder.insert_field(" MERGEFIELD Grade")

        # Create a data table that our mail merge will source data from.
        table = DataTable("ExamResults")
        table.columns.add("Name")
        table.columns.add("Grade")
        table.rows.add(["John Doe", "67"])
        table.rows.add(["Jane Doe", "81"])
        table.rows.add(["John Cardholder", "47"])
        table.rows.add(["Joe Bloggs", "75"])

        # We can use a data view to alter the mail merge data without making changes to the data table itself.
        view = DataView(table)
        view.sort = "Grade DESC"
        view.row_filter = "Grade >= 50"

        # Our data view sorts the entries in descending order along the "Grade" column
        # and filters out rows with values of less than 50 on that column.
        # Three out of the four rows fit those criteria so that the output document will contain three merge documents.
        doc.mail_merge.execute(view)

        doc.save(ARTIFACTS_DIR + "MailMerge.execute_data_view.docx")
        #ExEnd

        self.mail_merge_matches_data_table(view.to_table(), aw.Document(ARTIFACTS_DIR + "MailMerge.execute_data_view.docx"), True)

    #ExStart
    #ExFor:MailMerge.execute_with_regions(DataSet)
    #ExSummary:Shows how to execute a nested mail merge with two merge regions and two data tables.
    def test_execute_with_regions_nested(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Normally, MERGEFIELDs contain the name of a column of a mail merge data source.
        # Instead, we can use "TableStart:" and "TableEnd:" prefixes to begin/end a mail merge region.
        # Each region will belong to a table with a name that matches the string immediately after the prefix's colon.
        builder.insert_field(" MERGEFIELD TableStart:Customers")

        # This MERGEFIELD is inside the mail merge region of the "Customers" table.
        # When we execute the mail merge, this field will receive data from rows in a data source named "Customers".
        builder.write("Orders for ")
        builder.insert_field(" MERGEFIELD CustomerName")
        builder.write(":")

        # Create column headers for a table that will contain values from a second inner region.
        builder.start_table()
        builder.insert_cell()
        builder.write("Item")
        builder.insert_cell()
        builder.write("Quantity")
        builder.end_row()

        # Create a second mail merge region inside the outer region for a table named "Orders".
        # The "Orders" table has a many-to-one relationship with the "Customers" table on the "CustomerID" column.
        builder.insert_cell()
        builder.insert_field(" MERGEFIELD TableStart:Orders")
        builder.insert_field(" MERGEFIELD ItemName")
        builder.insert_cell()
        builder.insert_field(" MERGEFIELD Quantity")

        # End the inner region, and then end the outer region. The opening and closing of a mail merge region must
        # happen on the same row of a table.
        builder.insert_field(" MERGEFIELD TableEnd:Orders")
        builder.end_table()

        builder.insert_field(" MERGEFIELD TableEnd:Customers")

        # Create a dataset that contains the two tables with the required names and relationships.
        # Each merge document for each row of the "Customers" table of the outer merge region will perform its mail merge on the "Orders" table.
        # Each merge document will display all rows of the latter table whose "CustomerID" column values match the current "Customers" table row.
        customers_and_orders = ExMailMerge.create_data_set()
        doc.mail_merge.execute_with_regions(customers_and_orders)

        doc.save(ARTIFACTS_DIR + "MailMerge.execute_with_regions_nested.docx")
        self.mail_merge_matches_data_set(customers_and_orders, aw.Document(ARTIFACTS_DIR + "MailMerge.execute_with_regions_nested.docx"), False) #ExSkip

    @staticmethod
    def create_data_set() -> DataSet:
        """Generates a data set that has two data tables named "Customers" and "Orders", with a one-to-many relationship on the "CustomerID" column."""

        table_customers = DataTable("Customers")
        table_customers.columns.add("CustomerID")
        table_customers.columns.add("CustomerName")
        table_customers.rows.add([1, "John Doe"])
        table_customers.rows.add([2, "Jane Doe"])

        table_orders = DataTable("Orders")
        table_orders.columns.add("CustomerID")
        table_orders.columns.add("ItemName")
        table_orders.columns.add("Quantity")
        table_orders.rows.add([1, "Hawaiian", 2])
        table_orders.rows.add([2, "Pepperoni", 1])
        table_orders.rows.add([2, "Chicago", 1])

        data_set = DataSet()
        data_set.tables.add(table_customers)
        data_set.tables.add(table_orders)
        data_set.relations.add(table_customers.columns["CustomerID"], table_orders.columns["CustomerID"])

        return data_set

    #ExEnd

    def test_execute_with_regions_concurrent(self):

        #ExStart
        #ExFor:MailMerge.execute_with_regions(DataTable)
        #ExFor:MailMerge.execute_with_regions(DataView)
        #ExSummary:Shows how to use regions to execute two separate mail merges in one document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # If we want to perform two consecutive mail merges on one document while taking data from two tables
        # related to each other in any way, we can separate the mail merges with regions.
        # Normally, MERGEFIELDs contain the name of a column of a mail merge data source.
        # Instead, we can use "TableStart:" and "TableEnd:" prefixes to begin/end a mail merge region.
        # Each region will belong to a table with a name that matches the string immediately after the prefix's colon.
        # These regions are separate for unrelated data, while they can be nested for hierarchical data.
        builder.writeln("\tCities: ")
        builder.insert_field(" MERGEFIELD TableStart:Cities")
        builder.insert_field(" MERGEFIELD Name")
        builder.insert_field(" MERGEFIELD TableEnd:Cities")
        builder.insert_paragraph()

        # Both MERGEFIELDs refer to the same column name, but values for each will come from different data tables.
        builder.writeln("\tFruit: ")
        builder.insert_field(" MERGEFIELD TableStart:Fruit")
        builder.insert_field(" MERGEFIELD Name")
        builder.insert_field(" MERGEFIELD TableEnd:Fruit")

        # Create two unrelated data tables.
        table_cities = DataTable("Cities")
        table_cities.columns.add("Name")
        table_cities.rows.add(["Washington"])
        table_cities.rows.add(["London"])
        table_cities.rows.add(["New York"])

        table_fruit = DataTable("Fruit")
        table_fruit.columns.add("Name")
        table_fruit.rows.add(["Cherry"])
        table_fruit.rows.add(["Apple"])
        table_fruit.rows.add(["Watermelon"])
        table_fruit.rows.add(["Banana"])

        # We will need to run one mail merge per table. The first mail merge will populate the MERGEFIELDs
        # in the "Cities" range while leaving the fields the "Fruit" range unfilled.
        doc.mail_merge.execute_with_regions(table_cities)

        # Run a second merge for the "Fruit" table, while using a data view
        # to sort the rows in ascending order on the "Name" column before the merge.
        data_view = DataView(table_fruit)
        data_view.sort = "Name ASC"
        doc.mail_merge.execute_with_regions(data_view)

        doc.save(ARTIFACTS_DIR + "MailMerge.execute_with_regions_concurrent.docx")
        #ExEnd

        data_set = DataSet()

        data_set.tables.add(table_cities)
        data_set.tables.add(table_fruit)

        self.mail_merge_matches_data_set(data_set, aw.Document(ARTIFACTS_DIR + "MailMerge.execute_with_regions_concurrent.docx"), False)

    def test_mail_merge_region_info(self):

        #ExStart
        #ExFor:MailMerge.get_field_names_for_region(str)
        #ExFor:MailMerge.get_field_names_for_region(str,int)
        #ExFor:MailMerge.get_regions_by_name(str)
        #ExFor:MailMerge.region_end_tag
        #ExFor:MailMerge.region_start_tag
        #ExFor:MailMergeRegionInfo.parent_region
        #ExSummary:Shows how to create, list, and read mail merge regions.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # "TableStart" and "TableEnd" tags, which go inside MERGEFIELDs,
        # denote the strings that signify the starts and ends of mail merge regions.
        self.assertEqual("TableStart", doc.mail_merge.region_start_tag)
        self.assertEqual("TableEnd", doc.mail_merge.region_end_tag)

        # Use these tags to start and end a mail merge region named "MailMergeRegion1",
        # which will contain MERGEFIELDs for two columns.
        builder.insert_field(" MERGEFIELD TableStart:MailMergeRegion1")
        builder.insert_field(" MERGEFIELD Column1")
        builder.write(", ")
        builder.insert_field(" MERGEFIELD Column2")
        builder.insert_field(" MERGEFIELD TableEnd:MailMergeRegion1")

        # We can keep track of merge regions and their columns by looking at these collections.
        regions = doc.mail_merge.get_regions_by_name("MailMergeRegion1")

        self.assertEqual(1, regions.count)
        self.assertEqual("MailMergeRegion1", regions[0].name)

        merge_field_names = doc.mail_merge.get_field_names_for_region("MailMergeRegion1")

        self.assertEqual("Column1", merge_field_names[0])
        self.assertEqual("Column2", merge_field_names[1])

        # Insert a region with the same name inside the existing region, which will make it a parent.
        # Now a "Column2" field will be inside a new region.
        builder.move_to_field(regions[0].fields[1], False)
        builder.insert_field(" MERGEFIELD TableStart:MailMergeRegion1")
        builder.move_to_field(regions[0].fields[1], True)
        builder.insert_field(" MERGEFIELD TableEnd:MailMergeRegion1")

        # If we look up the name of duplicate regions using the "GetRegionsByName" method,
        # it will return all such regions in a collection.
        regions = doc.mail_merge.get_regions_by_name("MailMergeRegion1")

        self.assertEqual(2, regions.count)
        # Check that the second region now has a parent region.
        self.assertEqual("MailMergeRegion1", regions[1].parent_region.name)

        merge_field_names = doc.mail_merge.get_field_names_for_region("MailMergeRegion1", 1)

        self.assertEqual("Column2", merge_field_names[0])
        #ExEnd

    #ExStart
    #ExFor:MailMerge.merge_duplicate_regions
    #ExSummary:Shows how to work with duplicate mail merge regions.
    def test_merge_duplicate_regions(self):

        for merge_duplicate_regions in (True, False):
            with self.subTest(merge_duplicate_regions=merge_duplicate_regions):
                doc = ExMailMerge.create_source_doc_merge_duplicate_regions()
                data_table = ExMailMerge.create_source_table_merge_duplicate_regions()

                # If we set the "merge_duplicate_regions" property to "False", the mail merge will affect the first region,
                # while the MERGEFIELDs of the second one will be left in the pre-merge state.
                # To get both regions merged like that,
                # we would have to execute the mail merge twice on a table of the same name.
                # If we set the "merge_duplicate_regions" property to "True", the mail merge will affect both regions.
                doc.mail_merge.merge_duplicate_regions = merge_duplicate_regions

                doc.mail_merge.execute_with_regions(data_table)
                doc.save(ARTIFACTS_DIR + "MailMerge.merge_duplicate_regions.docx")
                self._test_merge_duplicate_regions(data_table, doc, merge_duplicate_regions) #ExSkip

    @staticmethod
    def create_source_doc_merge_duplicate_regions() -> aw.Document:
        """Returns a document that contains two duplicate mail merge regions (sharing the same name in the "TableStart/End" tags)."""

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_field(" MERGEFIELD TableStart:MergeRegion")
        builder.insert_field(" MERGEFIELD Column1")
        builder.insert_field(" MERGEFIELD TableEnd:MergeRegion")
        builder.insert_paragraph()

        builder.insert_field(" MERGEFIELD TableStart:MergeRegion")
        builder.insert_field(" MERGEFIELD Column2")
        builder.insert_field(" MERGEFIELD TableEnd:MergeRegion")

        return doc

    @staticmethod
    def create_source_table_merge_duplicate_regions() -> DataTable:
        """Creates a data table with one row and two columns."""

        data_table = DataTable("MergeRegion")
        data_table.columns.add("Column1")
        data_table.columns.add("Column2")
        data_table.rows.add(["Value 1", "Value 2"])

        return data_table

    #ExEnd

    def _test_merge_duplicate_regions(self, data_table: DataTable, doc: aw.Document, is_merge_duplicate_regions: bool):

        if is_merge_duplicate_regions:
            self.mail_merge_matches_data_table(data_table, doc, True)
        else:
            data_table.columns.remove("Column2")
            self.mail_merge_matches_data_table(data_table, doc, True)

    #ExStart
    #ExFor:MailMerge.preserve_unused_tags
    #ExFor:MailMerge.use_non_merge_fields
    #ExSummary:Shows how to preserve the appearance of alternative mail merge tags that go unused during a mail merge.
    def test_preserve_unused_tags(self):

        for preserve_unused_tags in (False, True):
            with self.subTest(preserve_unused_tags=preserve_unused_tags):
                doc = ExMailMerge.create_source_doc_with_alternative_merge_fields()
                data_table = ExMailMerge.create_source_table_preserve_unused_tags()

                # By default, a mail merge places data from each row of a table into MERGEFIELDs, which name columns in that table.
                # Our document has no such fields, but it does have plaintext tags enclosed by curly braces.
                # If we set the "preserve_unused_tags" flag to "True", we could treat these tags as MERGEFIELDs
                # to allow our mail merge to insert data from the data source at those tags.
                # If we set the "preserve_unused_tags" flag to "False",
                # the mail merge will convert these tags to MERGEFIELDs and leave them unfilled.
                doc.mail_merge.preserve_unused_tags = preserve_unused_tags
                doc.mail_merge.execute(data_table)

                doc.save(ARTIFACTS_DIR + "MailMerge.preserve_unused_tags.docx")

                # Our document has a tag for a column named "Column2", which does not exist in the table.
                # If we set the "preserve_unused_tags" flag to "False", then the mail merge will convert this tag into a MERGEFIELD.
                self.assertEqual(doc.get_text().contains("{{ Column2 }}"), preserve_unused_tags)

                if preserve_unused_tags:
                    self.assertEqual(0, len([f for f in doc.range.fields if f.type == aw.fields.field_type.FIELD_MERGE_FIELD]))
                else:
                    self.assertEqual(1, len([f for f in doc.range.fields if f.type == aw.fields.field_type.FIELD_MERGE_FIELD]))
                self.mail_merge_matches_data_table(data_table, doc, True) #ExSkip

    @staticmethod
    def create_source_doc_with_alternative_merge_fields() -> aw.Document:
        """Create a document and add two plaintext tags that may act as MERGEFIELDs during a mail merge."""

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("{{ Column1 }}")
        builder.writeln("{{ Column2 }}")

        # Our tags will register as destinations for mail merge data only if we set this to True.
        doc.mail_merge.use_non_merge_fields = True

        return doc

    @staticmethod
    def create_source_table_preserve_unused_tags() -> DataTable:
        """Create a simple data table with one column."""

        data_table = DataTable("MyTable")
        data_table.columns.add("Column1")
        data_table.rows.add(["Value1"])

        return data_table

    #ExEnd

    #ExStart
    #ExFor:MailMerge.merge_whole_document
    #ExSummary:Shows the relationship between mail merges with regions, and field updating.
    def test_merge_whole_document(self):

        for merge_whole_document in (False, True):
            with self.subTest(merge_whole_document=merge_whole_document):
                doc = ExMailMerge.create_source_doc_merge_whole_document()
                data_table = ExMailMerge.create_source_table_merge_whole_document()

                # If we set the "merge_whole_document" flag to "True",
                # the mail merge with regions will update every field in the document.
                # If we set the "merge_whole_document" flag to "False", the mail merge will only update fields
                # within the mail merge region whose name matches the name of the data source table.
                doc.mail_merge.merge_whole_document = merge_whole_document
                doc.mail_merge.execute_with_regions(data_table)

                # The mail merge will only update the QUOTE field outside of the mail merge region
                # if we set the "merge_whole_document" flag to "True".
                doc.save(ARTIFACTS_DIR + "MailMerge.merge_whole_document.docx")

                self.assertTrue(doc.get_text().contains("This QUOTE field is inside the \"MyTable\" merge region."))
                self.assertEqual(merge_whole_document,
                    doc.get_text().contains("This QUOTE field is outside of the \"MyTable\" merge region."))
                self.mail_merge_matches_data_table(data_table, doc, True) #ExSkip

    @staticmethod
    def create_source_doc_merge_whole_document() -> aw.Document:
        """Create a document with a mail merge region that belongs to a data source named "MyTable".
        Insert one QUOTE field inside this region, and one more outside it."""

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        field = builder.insert_field(aw.fields.field_type.FIELD_QUOTE, True).as_field_quote()
        field.text = "This QUOTE field is outside of the \"MyTable\" merge region."

        builder.insert_paragraph()
        builder.insert_field(" MERGEFIELD TableStart:MyTable")

        field = builder.insert_field(aw.fields.field_type.FIELD_QUOTE, True).as_field_quote()
        field.text = "This QUOTE field is inside the \"MyTable\" merge region."
        builder.insert_paragraph()

        builder.insert_field(" MERGEFIELD MyColumn")
        builder.insert_field(" MERGEFIELD TableEnd:MyTable")

        return doc

    @staticmethod
    def create_source_table_merge_whole_document() -> DataTable:
        """Create a data table that will be used in a mail merge."""

        data_table = DataTable("MyTable")
        data_table.columns.add("MyColumn")
        data_table.rows.add(["MyValue"])

        return data_table

    #ExEnd

    #ExStart
    #ExFor:MailMerge.use_whole_paragraph_as_region
    #ExSummary:Shows the relationship between mail merge regions and paragraphs.
    def test_use_whole_paragraph_as_region(self):

        for use_whole_paragraph_as_region in (False, True):
            with self.subTest(use_whole_paragraph_as_region=use_whole_paragraph_as_region):
                doc = ExMailMerge.create_source_doc_with_nested_merge_regions()
                data_table = ExMailMerge.create_source_table_data_table_for_one_region()

                # By default, a paragraph can belong to no more than one mail merge region.
                # The contents of our document do not meet these criteria.
                # If we set the "use_whole_paragraph_as_region" flag to "True",
                # running a mail merge on this document will throw an exception.
                # If we set the "use_whole_paragraph_as_region" flag to "False",
                # we will be able to execute a mail merge on this document.
                doc.mail_merge.use_whole_paragraph_as_region = use_whole_paragraph_as_region

                if use_whole_paragraph_as_region:
                    with self.assertRaises(Exception):
                        doc.mail_merge.execute_with_regions(data_table)
                else:
                    doc.mail_merge.execute_with_regions(data_table)

                # The mail merge populates our first region while leaving the second region unused
                # since it is the region that breaks the rule.
                doc.save(ARTIFACTS_DIR + "MailMerge.use_whole_paragraph_as_region.docx")
                if not use_whole_paragraph_as_region: #ExSkip
                    self.mail_merge_matches_data_table(data_table, aw.Document(ARTIFACTS_DIR + "MailMerge.use_whole_paragraph_as_region.docx"), True) #ExSkip

    @staticmethod
    def create_source_doc_with_nested_merge_regions() -> aw.Document:
        """Create a document with two mail merge regions sharing one paragraph."""

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Region 1: ")
        builder.insert_field(" MERGEFIELD TableStart:MyTable")
        builder.insert_field(" MERGEFIELD Column1")
        builder.write(", ")
        builder.insert_field(" MERGEFIELD Column2")
        builder.insert_field(" MERGEFIELD TableEnd:MyTable")

        builder.write(", Region 2: ")
        builder.insert_field(" MERGEFIELD TableStart:MyOtherTable")
        builder.insert_field(" MERGEFIELD TableEnd:MyOtherTable")

        return doc

    @staticmethod
    def create_source_table_data_table_for_one_region() -> DataTable:
        """Create a data table that can populate one region during a mail merge."""

        data_table = DataTable("MyTable")
        data_table.columns.add("Column1")
        data_table.columns.add("Column2")
        data_table.rows.add(["Value 1", "Value 2"])

        return data_table

    #ExEnd

    def test_trim_white_spaces(self):

        for trim_whitespaces in (False, True):
            with self.subTest(trim_whitespaces=trim_whitespaces):
                #ExStart
                #ExFor:MailMerge.trim_whitespaces
                #ExSummary:Shows how to trim whitespaces from values of a data source while executing a mail merge.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.insert_field("MERGEFIELD myMergeField", None)

                doc.mail_merge.trim_whitespaces = trim_whitespaces
                doc.mail_merge.execute(["myMergeField"], ["\t hello world! "])

                self.assertEqual("hello world!\f" if trim_whitespaces else "\t hello world! \f", doc.get_text())
                #ExEnd

    def test_mail_merge_get_field_names(self):

        #ExStart
        #ExFor:MailMerge.get_field_names
        #ExSummary:Shows how to get names of all merge fields in a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_field(" MERGEFIELD FirstName ")
        builder.write(" ")
        builder.insert_field(" MERGEFIELD LastName ")
        builder.insert_paragraph()
        builder.insert_field(" MERGEFIELD City ")

        data_table = DataTable("MyTable")
        data_table.columns.add("FirstName")
        data_table.columns.add("LastName")
        data_table.columns.add("City")
        data_table.rows.add(["John", "Doe", "New York"])
        data_table.rows.add(["Joe", "Bloggs", "Washington"])

        # For every MERGEFIELD name in the document, ensure that the data table contains a column
        # with the same name, and then execute the mail merge.
        field_names = doc.mail_merge.get_field_names()

        self.assertEqual(3, len(field_names))

        for field_name in field_names:
            self.assertTrue(data_table.columns.contains(field_name))

        doc.mail_merge.execute(data_table)
        #ExEnd

        self.mail_merge_matches_data_table(data_table, doc, True)

    def test_delete_fields(self):

        #ExStart
        #ExFor:MailMerge.delete_fields
        #ExSummary:Shows how to delete all MERGEFIELDs from a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Dear ")
        builder.insert_field(" MERGEFIELD FirstName ")
        builder.write(" ")
        builder.insert_field(" MERGEFIELD LastName ")
        builder.writeln(",")
        builder.writeln("Greetings!")

        self.assertEqual(
            "Dear \u0013 MERGEFIELD FirstName \u0014«FirstName»\u0015 \u0013 MERGEFIELD LastName \u0014«LastName»\u0015,\rGreetings!",
            doc.get_text().strip())

        doc.mail_merge.delete_fields()

        self.assertEqual("Dear  ,\rGreetings!", doc.get_text().strip())
        #ExEnd

    def test_remove_unused_fields(self):

        for mail_merge_cleanup_options in (aw.mailmerging.MailMergeCleanupOptions.NONE,
                                           aw.mailmerging.MailMergeCleanupOptions.REMOVE_CONTAINING_FIELDS,
                                           aw.mailmerging.MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS,
                                           aw.mailmerging.MailMergeCleanupOptions.REMOVE_EMPTY_TABLE_ROWS,
                                           aw.mailmerging.MailMergeCleanupOptions.REMOVE_STATIC_FIELDS,
                                           aw.mailmerging.MailMergeCleanupOptions.REMOVE_UNUSED_FIELDS,
                                           aw.mailmerging.MailMergeCleanupOptions.REMOVE_UNUSED_REGIONS):
            with self.subTest(mail_merge_cleanup_option=mail_merge_cleanup_options):
                #ExStart
                #ExFor:MailMerge.cleanup_options
                #ExFor:MailMergeCleanupOptions
                #ExSummary:Shows how to automatically remove MERGEFIELDs that go unused during mail merge.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Create a document with MERGEFIELDs for three columns of a mail merge data source table,
                # and then create a table with only two columns whose names match our MERGEFIELDs.
                builder.insert_field(" MERGEFIELD FirstName ")
                builder.write(" ")
                builder.insert_field(" MERGEFIELD LastName ")
                builder.insert_paragraph()
                builder.insert_field(" MERGEFIELD City ")

                data_table = DataTable("MyTable")
                data_table.columns.add("FirstName")
                data_table.columns.add("LastName")
                data_table.rows.add(["John", "Doe"])
                data_table.rows.add(["Joe", "Bloggs"])

                # Our third MERGEFIELD references a "City" column, which does not exist in our data source.
                # The mail merge will leave fields such as this intact in their pre-merge state.
                # Setting the "cleanup_options" property to "REMOVE_UNUSED_FIELDS" will remove any MERGEFIELDs
                # that go unused during a mail merge to clean up the merge documents.
                doc.mail_merge.cleanup_options = mail_merge_cleanup_options
                doc.mail_merge.execute(data_table)

                if mail_merge_cleanup_options in (aw.mailmerging.MailMergeCleanupOptions.REMOVE_UNUSED_FIELDS, aw.mailmerging.MailMergeCleanupOptions.REMOVE_STATIC_FIELDS):
                    self.assertEqual(0, doc.range.fields.count)
                else:
                    self.assertEqual(2, doc.range.fields.count)
                #ExEnd

                self.mail_merge_matches_data_table(data_table, doc, True)

    def test_remove_empty_paragraphs(self):

        for mail_merge_cleanup_options in (aw.mailmerging.MailMergeCleanupOptions.NONE,
                                           aw.mailmerging.MailMergeCleanupOptions.REMOVE_CONTAINING_FIELDS,
                                           aw.mailmerging.MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS,
                                           aw.mailmerging.MailMergeCleanupOptions.REMOVE_EMPTY_TABLE_ROWS,
                                           aw.mailmerging.MailMergeCleanupOptions.REMOVE_STATIC_FIELDS,
                                           aw.mailmerging.MailMergeCleanupOptions.REMOVE_UNUSED_FIELDS,
                                           aw.mailmerging.MailMergeCleanupOptions.REMOVE_UNUSED_REGIONS):
            with self.subTest(mail_merge_cleanup_options=mail_merge_cleanup_options):
                #ExStart
                #ExFor:MailMerge.cleanup_options
                #ExFor:MailMergeCleanupOptions
                #ExSummary:Shows how to remove empty paragraphs that a mail merge may create from the merge output document.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.insert_field(" MERGEFIELD TableStart:MyTable")
                builder.insert_field(" MERGEFIELD FirstName ")
                builder.write(" ")
                builder.insert_field(" MERGEFIELD LastName ")
                builder.insert_field(" MERGEFIELD TableEnd:MyTable")

                data_table = DataTable("MyTable")
                data_table.columns.add("FirstName")
                data_table.columns.add("LastName")
                data_table.rows.add(["John", "Doe"])
                data_table.rows.add(["", ""])
                data_table.rows.add(["Jane", "Doe"])

                doc.mail_merge.cleanup_options = mail_merge_cleanup_options
                doc.mail_merge.execute_with_regions(data_table)

                if doc.mail_merge.cleanup_options == aw.mailmerging.MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS:
                    self.assertEqual(
                        "John Doe\r" +
                        "Jane Doe", doc.get_text().strip())
                else:
                    self.assertEqual(
                        "John Doe\r" +
                        " \r" +
                        "Jane Doe", doc.get_text().strip())
                #ExEnd

                self.mail_merge_matches_data_table(data_table, doc, False)

    # Ignore("WORDSNET-17733")
    def test_remove_colon_between_empty_merge_fields(self):

        parameters = [
            ("!", False, ""),
            (", ", False, ""),
            (" . ", False, ""),
            (" :", False, ""),
            ("  ; ", False, ""),
            (" ?  ", False, ""),
            ("  ¡  ", False, ""),
            ("  ¿  ", False, ""),
            ("!", True, "!\f"),
            (", ", True, ", \f"),
            (" . ", True, " . \f"),
            (" :", True, " :\f"),
            ("  ; ", True, "  ; \f"),
            (" ?  ", True, " ?  \f"),
            ("  ¡  ", True, "  ¡  \f"),
            ("  ¿  ", True, "  ¿  \f"),
            ]

        for punctuation_mark, cleanup_paragraphs_with_punctuation_marks, result_text in parameters:
            with self.subTest(punctuation_mark=punctuation_mark,
                              cleanup_paragraphs_with_punctuation_marks=cleanup_paragraphs_with_punctuation_marks,
                              result_text=result_text):
                #ExStart
                #ExFor:MailMerge.cleanup_paragraphs_with_punctuation_marks
                #ExSummary:Shows how to remove paragraphs with punctuation marks after a mail merge operation.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                merge_field_option1 = builder.insert_field("MERGEFIELD", "Option_1").as_field_merge_field()
                merge_field_option1.field_name = "Option_1"

                builder.write(punctuation_mark)

                merge_field_option2 = builder.insert_field("MERGEFIELD", "Option_2").as_field_merge_field()
                merge_field_option2.field_name = "Option_2"

                # Configure the "cleanup_options" property to remove any empty paragraphs that this mail merge would create.
                doc.mail_merge.cleanup_options = aw.mailmerging.MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS

                # Setting the "cleanup_paragraphs_with_punctuation_marks" property to "True" will also count paragraphs
                # with punctuation marks as empty and will get the mail merge operation to remove them as well.
                # Setting the "cleanup_paragraphs_with_punctuation_marks" property to "False"
                # will remove empty paragraphs, but not ones with punctuation marks.
                # This is a list of punctuation marks that this property concerns: "!", ",", ".", ":", ";", "?", "¡", "¿".
                doc.mail_merge.cleanup_paragraphs_with_punctuation_marks = cleanup_paragraphs_with_punctuation_marks

                doc.mail_merge.execute(["Option_1", "Option_2"], [None, None])

                doc.save(ARTIFACTS_DIR + "MailMerge.remove_colon_between_empty_merge_fields.docx")
                #ExEnd

                self.assertEqual(result_text, doc.get_text())

    #ExStart
    #ExFor:MailMerge.mapped_data_fields
    #ExFor:MappedDataFieldCollection
    #ExFor:MappedDataFieldCollection.add
    #ExFor:MappedDataFieldCollection.clear
    #ExFor:MappedDataFieldCollection.contains_key(str)
    #ExFor:MappedDataFieldCollection.contains_value(str)
    #ExFor:MappedDataFieldCollection.count
    #ExFor:MappedDataFieldCollection.__iter__
    #ExFor:MappedDataFieldCollection.__getitem__(str)
    #ExFor:MappedDataFieldCollection.remove(str)
    #ExSummary:Shows how to map data columns and MERGEFIELDs with different names so the data is transferred between them during a mail merge.
    def test_mapped_data_field_collection(self):

        doc = ExMailMerge.create_source_doc_mapped_data_fields()
        data_table = ExMailMerge.create_source_table_mapped_fata_fields()

        # The table has a column named "Column2", but there are no MERGEFIELDs with that name.
        # Also, we have a MERGEFIELD named "Column3", but the data source does not have a column with that name.
        # If data from "Column2" is suitable for the "Column3" MERGEFIELD,
        # we can map that column name to the MERGEFIELD in the "MappedDataFields" key/value pair.
        mapped_data_fields = doc.mail_merge.mapped_data_fields

        # We can link a data source column name to a MERGEFIELD name like this.
        mapped_data_fields.add("MergeFieldName", "DataSourceColumnName")

        # Link the data source column named "Column2" to MERGEFIELDs named "Column3".
        mapped_data_fields.add("Column3", "Column2")

        # The MERGEFIELD name is the "key" to the respective data source column name "value".
        self.assertEqual("DataSourceColumnName", mapped_data_fields["MergeFieldName"])
        self.assertTrue(mapped_data_fields.contains_key("MergeFieldName"))
        self.assertTrue(mapped_data_fields.contains_value("DataSourceColumnName"))

        # Now if we run this mail merge, the "Column3" MERGEFIELDs will take data from "Column2" of the table.
        doc.mail_merge.execute(data_table)

        doc.save(ARTIFACTS_DIR + "MailMerge.mapped_data_field_collection.docx")

        # We can iterate over the elements in this collection.
        self.assertEqual(2, mapped_data_fields.count)

        for field in mapped_data_fields:
            print(f"Column named {field.value} is mapped to MERGEFIELDs named {field.key}")

        # We can also remove elements from the collection.
        mapped_data_fields.remove("MergeFieldName")

        self.assertFalse(mapped_data_fields.contains_key("MergeFieldName"))
        self.assertFalse(mapped_data_fields.contains_value("DataSourceColumnName"))

        mapped_data_fields.clear()

        self.assertEqual(0, mapped_data_fields.count)
        self.mail_merge_matches_data_table(data_table, aw.Document(ARTIFACTS_DIR + "MailMerge.mapped_data_field_collection.docx"), True) #ExSkip

    @staticmethod
    def create_source_doc_mapped_data_fields() -> aw.Document:
        """Create a document with 2 MERGEFIELDs, one of which does not have a
        corresponding column in the data table from the method below."""

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_field(" MERGEFIELD Column1")
        builder.write(", ")
        builder.insert_field(" MERGEFIELD Column3")

        return doc

    @staticmethod
    def create_source_table_mapped_fata_fields() -> DataTable:
        """Create a data table with 2 columns, one of which does not have a
        corresponding MERGEFIELD in the source document from the method above."""

        data_table = DataTable("MyTable")
        data_table.columns.add("Column1")
        data_table.columns.add("Column2")
        data_table.rows.add(["Value1", "Value2"])

        return data_table

    #ExEnd

    def test_get_field_names(self):

        #ExStart
        #ExFor:FieldAddressBlock
        #ExFor:FieldAddressBlock.get_field_names
        #ExSummary:Shows how to get mail merge field names used by a field.
        doc = aw.Document(MY_DIR + "Field sample - ADDRESSBLOCK.docx")

        address_fields_expect = [
            "Company", "First Name", "Middle Name", "Last Name", "Suffix", "Address 1", "City", "State",
            "Country or Region", "Postal Code"
            ]

        address_block_field = doc.range.fields[0].as_field_address_block()
        address_block_field_names = address_block_field.get_field_names()
        #ExEnd

        self.assertEqual(address_fields_expect, address_block_field_names)

        greeting_fields_expect = ["Courtesy Title", "Last Name"]

        greeting_line_field = doc.range.fields[1].as_field_greeting_line()
        greeting_line_field_names = greeting_line_field.get_field_names()

        self.assertEqual(greeting_fields_expect, greeting_line_field_names)

    def test_mustache_template_syntax_true(self):
        """Without TestCaseSource/TestCase because of some strange behavior when using long data."""

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.write("{{ testfield1 }}")
        builder.write("{{ testfield2 }}")
        builder.write("{{ testfield3 }}")

        doc.mail_merge.use_non_merge_fields = True
        doc.mail_merge.preserve_unused_tags = True

        table = DataTable("Test")
        table.columns.add("testfield2")
        table.rows.add("value 1")

        doc.mail_merge.execute(table)

        para_text = DocumentHelper.get_paragraph_text(doc, 0)

        self.assertEqual("{{ testfield1 }}value 1{{ testfield3 }}\f", para_text)

    def test_mustache_template_syntax_false(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.write("{{ testfield1 }}")
        builder.write("{{ testfield2 }}")
        builder.write("{{ testfield3 }}")

        doc.mail_merge.use_non_merge_fields = True
        doc.mail_merge.preserve_unused_tags = False

        table = DataTable("Test")
        table.columns.add("testfield2")
        table.rows.add("value 1")

        doc.mail_merge.execute(table)

        para_text = DocumentHelper.get_paragraph_text(doc, 0)

        self.assertEqual("\u0013MERGEFIELD \"testfield1\"\u0014«testfield1»\u0015value 1\u0013MERGEFIELD \"testfield3\"\u0014«testfield3»\u0015\f", para_text)

    def test_test_mail_merge_get_regions_hierarchy(self):

        #ExStart
        #ExFor:MailMerge.get_regions_hierarchy
        #ExFor:MailMergeRegionInfo
        #ExFor:MailMergeRegionInfo.regions
        #ExFor:MailMergeRegionInfo.name
        #ExFor:MailMergeRegionInfo.fields
        #ExFor:MailMergeRegionInfo.start_field
        #ExFor:MailMergeRegionInfo.end_field
        #ExFor:MailMergeRegionInfo.level
        #ExSummary:Shows how to verify mail merge regions.
        doc = aw.Document(MY_DIR + "Mail merge regions.docx")

        # Returns a full hierarchy of merge regions that contain MERGEFIELDs available in the document.
        region_info = doc.mail_merge.get_regions_hierarchy()

        # Get top regions in the document.
        top_regions = region_info.regions

        self.assertEqual(2, top_regions.count)
        self.assertEqual("Region1", top_regions[0].name)
        self.assertEqual("Region2", top_regions[1].name)
        self.assertEqual(1, top_regions[0].level)
        self.assertEqual(1, top_regions[1].level)

        # Get nested region in first top region.
        nested_regions = top_regions[0].regions

        self.assertEqual(2, nested_regions.count)
        self.assertEqual("NestedRegion1", nested_regions[0].name)
        self.assertEqual("NestedRegion2", nested_regions[1].name)
        self.assertEqual(2, nested_regions[0].level)
        self.assertEqual(2, nested_regions[1].level)

        # Get list of fields inside the first top region.
        field_list = top_regions[0].fields

        self.assertEqual(4, field_list.count)

        start_field_merge_field = nested_regions[0].start_field

        self.assertEqual("TableStart:NestedRegion1", start_field_merge_field.field_name)

        end_field_merge_field = nested_regions[0].end_field

        self.assertEqual("TableEnd:NestedRegion1", end_field_merge_field.field_name)
        #ExEnd

    #ExStart
    #ExFor:MailMerge.mail_merge_callback
    #ExFor:IMailMergeCallback
    #ExFor:IMailMergeCallback.tags_replaced
    #ExSummary:Shows how to define custom logic for handling events during mail merge.
    def test_callback(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert two mail merge tags referencing two columns in a data source.
        builder.write("{{FirstName}}")
        builder.write("{{LastName}}")

        # Create a data source that only contains one of the columns that our merge tags reference.
        table = DataTable("Test")
        table.columns.add("FirstName")
        table.rows.add("John")
        table.rows.add("Jane")

        # Configure our mail merge to use alternative mail merge tags.
        doc.mail_merge.use_non_merge_fields = True

        # Then, ensure that the mail merge will convert tags, such as our "LastName" tag,
        # into MERGEFIELDs in the merge documents.
        doc.mail_merge.preserve_unused_tags = False

        counter = ExMailMerge.MailMergeTagReplacementCounter()
        doc.mail_merge.mail_merge_callback = counter
        doc.mail_merge.execute(table)

        self.assertEqual(1, counter.tags_replaced_count)

    class MailMergeTagReplacementCounter(aw.mailmerging.IMailMergeCallback):
        """Counts the number of times a mail merge replaces mail merge tags that it could not fill with data with MERGEFIELDs."""

        def __init__(self):
            self.tags_replaced_count = 0

        def test_tags_replaced(self):

            self.tags_replaced_count += 1

    #ExEnd

    def test_get_regions_by_name(self):

        doc = aw.Document(MY_DIR + "Mail merge regions.docx")

        regions = doc.mail_merge.get_regions_by_name("Region1")
        self.assertEqual(1, doc.mail_merge.get_regions_by_name("Region1").count)
        for region in regions:
            self.assertEqual("Region1", region.name)

        regions = doc.mail_merge.get_regions_by_name("Region2")
        self.assertEqual(1, doc.mail_merge.get_regions_by_name("Region2").count)
        for region in regions:
            self.assertEqual("Region2", region.name)

        regions = doc.mail_merge.get_regions_by_name("NestedRegion1")
        self.assertEqual(2, doc.mail_merge.get_regions_by_name("NestedRegion1").count)
        for region in regions:
            self.assertEqual("NestedRegion1", region.name)

    def test_cleanup_options(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.start_table()
        builder.insert_cell()
        builder.insert_field(" MERGEFIELD  TableStart:StudentCourse ")
        builder.insert_cell()
        builder.insert_field(" MERGEFIELD  CourseName ")
        builder.insert_cell()
        builder.insert_field(" MERGEFIELD  TableEnd:StudentCourse ")
        builder.end_table()

        data = ExMailMerge.get_data_table()

        doc.mail_merge.cleanup_options = aw.mailmerging.mail_merge_cleanup_options.REMOVE_EMPTY_TABLE_ROWS
        doc.mail_merge.execute_with_regions(data)

        doc.save(ARTIFACTS_DIR + "MailMerge.cleanup_options.docx")

        self.assertTrue(DocumentHelper.compare_docs(ARTIFACTS_DIR + "MailMerge.cleanup_options.docx", GOLDS_DIR + "MailMerge.CleanupOptions Gold.docx"))

    @staticmethod
    def get_data_table() -> DataTable:
        """Return a data table filled with sample data."""

        data_table = DataTable("StudentCourse")
        data_table.columns.add("CourseName")

        data_row_empty = data_table.new_row()
        data_table.rows.add(data_row_empty)
        data_row_empty[0] = ""

        for i in range(10):
            datarow = data_table.new_row()
            data_table.rows.add(datarow)
            datarow[0] = "Course " + i

        return data_table

    def test_unconditional_merge_fields_and_regions(self):

        for count_all_merge_fields in (False, True):
            with self.subTest(count_all_merge_fields=count_all_merge_fields):
                #ExStart
                #ExFor:MailMerge.unconditional_merge_fields_and_regions
                #ExSummary:Shows how to merge fields or regions regardless of the parent IF field's condition.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Insert a MERGEFIELD nested inside an IF field.
                # Since the IF field statement is False, it will not display the result of the MERGEFIELD.
                # The MERGEFIELD will also not receive any data during a mail merge.
                field_if = builder.insert_field(" IF 1 = 2 ").as_field_if()
                builder.move_to(field_if.separator)
                builder.insert_field(" MERGEFIELD  FullName ")

                # If we set the "unconditional_merge_fields_and_regions" flag to "True",
                # our mail merge will insert data into non-displayed fields such as our MERGEFIELD as well as all others.
                # If we set the "unconditional_merge_fields_and_regions" flag to "False",
                # our mail merge will not insert data into MERGEFIELDs hidden by IF fields with false statements.
                doc.mail_merge.unconditional_merge_fields_and_regions = count_all_merge_fields

                data_table = DataTable()
                data_table.columns.add("FullName")
                data_table.rows.add("James Bond")

                doc.mail_merge.execute(data_table)

                doc.save(ARTIFACTS_DIR + "MailMerge.unconditional_merge_fields_and_regions.docx")

                self.assertEqual(
                    "\u0013 IF 1 = 2 \"James Bond\"\u0014\u0015" if count_all_merge_fields else "\u0013 IF 1 = 2 \u0013 MERGEFIELD  FullName \u0014«FullName»\u0015\u0014\u0015",
                    doc.get_text().strip())
                #ExEnd

    def test_retain_first_section_start(self):

        parameters = [
            (True, aw.SectionStart.CONTINUOUS, aw.SectionStart.CONTINUOUS),
            (True, aw.SectionStart.NEW_COLUMN, aw.SectionStart.NEW_COLUMN),
            (True, aw.SectionStart.NEW_PAGE, aw.SectionStart.NEW_PAGE),
            (True, aw.SectionStart.EVEN_PAGE, aw.SectionStart.EVEN_PAGE),
            (True, aw.SectionStart.ODD_PAGE, aw.SectionStart.ODD_PAGE),
            (False, aw.SectionStart.CONTINUOUS, aw.SectionStart.NEW_PAGE),
            (False, aw.SectionStart.NEW_COLUMN, aw.SectionStart.NEW_PAGE),
            (False, aw.SectionStart.NEW_PAGE, aw.SectionStart.NEW_PAGE),
            (False, aw.SectionStart.EVEN_PAGE, aw.SectionStart.EVEN_PAGE),
            (False, aw.SectionStart.ODD_PAGE, aw.SectionStart.ODD_PAGE),
            ]

        for is_retain_first_section_start, section_start, expected in parameters:
            with self.subTest(is_retain_first_section_start=is_retain_first_section_start,
                              section_start=section_start,
                              expected=expected):
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.insert_field(" MERGEFIELD  FullName ")

                doc.first_section.page_setup.section_start = section_start
                doc.mail_merge.retainfirst_section_start = is_retain_first_section_start

                data_table = DataTable()
                data_table.columns.add("FullName")
                data_table.rows.add("James Bond")

                doc.mail_merge.execute(data_table)

                for section in doc.sections:
                    self.assertEqual(expected, section.page_setup.section_start)

    def test_mail_merge_settings(self):

        #ExStart
        #ExFor:Document.mail_merge_settings
        #ExFor:MailMergeCheckErrors
        #ExFor:MailMergeDataType
        #ExFor:MailMergeDestination
        #ExFor:MailMergeMainDocumentType
        #ExFor:MailMergeSettings
        #ExFor:MailMergeSettings.check_errors
        #ExFor:MailMergeSettings.clone
        #ExFor:MailMergeSettings.destination
        #ExFor:MailMergeSettings.data_type
        #ExFor:MailMergeSettings.do_not_supress_blank_lines
        #ExFor:MailMergeSettings.link_to_query
        #ExFor:MailMergeSettings.main_document_type
        #ExFor:MailMergeSettings.odso
        #ExFor:MailMergeSettings.query
        #ExFor:MailMergeSettings.view_merged_data
        #ExFor:Odso
        #ExFor:Odso.clone
        #ExFor:Odso.column_delimiter
        #ExFor:Odso.data_source
        #ExFor:Odso.data_source_type
        #ExFor:Odso.first_row_contains_column_names
        #ExFor:OdsoDataSourceType
        #ExSummary:Shows how to execute a mail merge with data from an Office Data Source Object.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Dear ")
        builder.insert_field("MERGEFIELD FirstName", "<FirstName>")
        builder.write(" ")
        builder.insert_field("MERGEFIELD LastName", "<LastName>")
        builder.writeln(": ")
        builder.insert_field("MERGEFIELD Message", "<Message>")

        # Create a data source in the form of an ASCII file, with the "|" character
        # acting as the delimiter that separates columns. The first line contains the three columns' names,
        # and each subsequent line is a row with their respective values.
        lines = [
            "FirstName|LastName|Message",
            "John|Doe|Hello! This message was created with Aspose Words mail merge."
            ]
        data_src_filename = ARTIFACTS_DIR + "MailMerge.mail_merge_settings.data_source.txt"

        with open(data_src_filename, "wt") as file:
            file.writelines(lines)

        settings = doc.mail_merge_settings
        settings.main_document_type = aw.settings.MailMergeMainDocumentType.MAILING_LABELS
        settings.check_errors = aw.settings.MailMergeCheckErrors.SIMULATE
        settings.data_type = aw.settings.MailMergeDataType.NATIVE
        settings.data_source = data_src_filename
        settings.query = "SELECT * FROM " + doc.mail_merge_settings.data_source
        settings.link_to_query = True
        settings.view_merged_data = True

        self.assertEqual(aw.settings.MailMergeDestination.DEFAULT, settings.destination)
        self.assertFalse(settings.do_not_supress_blank_lines)

        odso = settings.odso
        odso.data_source = data_src_filename
        odso.data_source_type = aw.settings.OdsoDataSourceType.TEXT
        odso.column_delimiter = '|'
        odso.first_row_contains_column_names = True

        #Assert.are_not_same(odso, odso.clone())
        #Assert.are_not_same(settings, settings.clone())

        # Opening this document in Microsoft Word will execute the mail merge before displaying the contents.
        doc.save(ARTIFACTS_DIR + "MailMerge.mail_merge_settings.docx")
        #ExEnd

        settings = aw.Document(ARTIFACTS_DIR + "MailMerge.mail_merge_settings.docx").mail_merge_settings

        self.assertEqual(aw.settings.MailMergeMainDocumentType.MAILING_LABELS, settings.main_document_type)
        self.assertEqual(aw.settings.MailMergeCheckErrors.SIMULATE, settings.check_errors)
        self.assertEqual(aw.settings.MailMergeDataType.NATIVE, settings.data_type)
        self.assertEqual(ARTIFACTS_DIR + "MailMerge.mail_merge_settings.data_source.txt", settings.data_source)
        self.assertEqual("SELECT * FROM " + doc.mail_merge_settings.data_source, settings.query)
        self.assertTrue(settings.link_to_query)
        self.assertTrue(settings.view_merged_data)

        odso = settings.odso
        self.assertEqual(ARTIFACTS_DIR + "MailMerge.mail_merge_settings.data_source.txt", odso.data_source)
        self.assertEqual(aw.settings.OdsoDataSourceType.TEXT, odso.data_source_type)
        self.assertEqual('|', odso.column_delimiter)
        self.assertTrue(odso.first_row_contains_column_names)

    def test_odso_email(self):

        #ExStart
        #ExFor:MailMergeSettings.active_record
        #ExFor:MailMergeSettings.address_field_name
        #ExFor:MailMergeSettings.connect_string
        #ExFor:MailMergeSettings.mail_as_attachment
        #ExFor:MailMergeSettings.mail_subject
        #ExFor:MailMergeSettings.clear
        #ExFor:Odso.table_name
        #ExFor:Odso.udl_connect_string
        #ExSummary:Shows how to execute a mail merge while connecting to an external data source.
        doc = aw.Document(MY_DIR + "Odso data.docx")
        self._test_odso_email(doc) #ExSkip
        settings = doc.mail_merge_settings

        print(f"Connection string:\n\t{settings.connect_string}")
        print(f"Mail merge docs as attachment:\n\t{settings.mail_as_attachment}")
        print(f"Mail merge doc e-mail subject:\n\t{settings.mail_subject}")
        print(f"Column that contains e-mail addresses:\n\t{settings.address_field_name}")
        print(f"Active record:\n\t{settings.active_record}")

        odso = settings.odso

        print(f"File will connect to data source located in:\n\t\"{odso.data_source}\"")
        print(f"Source type:\n\t{odso.data_source_type}")
        print(f"UDL connection string:\n\t{odso.udl_connect_string}")
        print(f"Table:\n\t{odso.table_name}")
        print(f"Query:\n\t{doc.mail_merge_settings.query}")

        # We can reset these settings by clearing them. Once we do that and save the document,
        # Microsoft Word will no longer execute a mail merge when we use it to load the document.
        settings.clear()

        doc.save(ARTIFACTS_DIR + "MailMerge.odso_email.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "MailMerge.odso_email.docx")
        self.assertEqual("", doc.mail_merge_settings.connect_string)

    def _test_odso_email(self, doc: aw.Document):

        settings = doc.mail_merge_settings

        self.assertFalse(settings.mail_as_attachment)
        self.assertEqual("test subject", settings.mail_subject)
        self.assertEqual("Email_Address", settings.address_field_name)
        self.assertEqual(66, settings.active_record)
        self.assertEqual("SELECT * FROM `Contacts` ", settings.query)

        odso = settings.odso

        self.assertEqual(settings.connect_string, odso.udl_connect_string)
        self.assertEqual("Personal Folders|", odso.data_source)
        self.assertEqual(aw.settings.OdsoDataSourceType.EMAIL, odso.data_source_type)
        self.assertEqual("Contacts", odso.table_name)

    def test_mailing_label_merge(self):

        #ExStart
        #ExFor:MailMergeSettings.data_source
        #ExFor:MailMergeSettings.header_source
        #ExSummary:Shows how to construct a data source for a mail merge from a header source and a data source.
        # Create a mailing label merge header file, which will consist of a table with one row.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_table()
        builder.insert_cell()
        builder.write("FirstName")
        builder.insert_cell()
        builder.write("LastName")
        builder.end_table()

        doc.save(ARTIFACTS_DIR + "MailMerge.mailing_label_merge.header.docx")

        # Create a mailing label merge data file consisting of a table with one row
        # and the same number of columns as the header document's table.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_table()
        builder.insert_cell()
        builder.write("John")
        builder.insert_cell()
        builder.write("Doe")
        builder.end_table()

        doc.save(ARTIFACTS_DIR + "MailMerge.mailing_label_merge.data.docx")

        # Create a merge destination document with MERGEFIELDS with names that
        # match the column names in the merge header file table.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Dear ")
        builder.insert_field("MERGEFIELD FirstName", "<FirstName>")
        builder.write(" ")
        builder.insert_field("MERGEFIELD LastName", "<LastName>")

        settings = doc.mail_merge_settings

        # Construct a data source for our mail merge by specifying two document filenames.
        # The header source will name the columns of the data source table.
        settings.header_source = ARTIFACTS_DIR + "MailMerge.mailing_label_merge.header.docx"

        # The data source will provide rows of data for all the columns in the header document table.
        settings.data_source = ARTIFACTS_DIR + "MailMerge.mailing_label_merge.data.docx"

        # Configure a mailing label type mail merge, which Microsoft Word will execute
        # as soon as we use it to load the output document.
        settings.query = "SELECT * FROM " + settings.data_source
        settings.main_document_type = aw.settings.MailMergeMainDocumentType.MAILING_LABELS
        settings.data_type = aw.settings.MailMergeDataType.TEXT_FILE
        settings.link_to_query = True
        settings.view_merged_data = True

        doc.save(ARTIFACTS_DIR + "MailMerge.mailing_label_merge.docx")
        #ExEnd

        self.assertEqual(
            "FirstName\aLastName\a\a",
            aw.Document(ARTIFACTS_DIR + "MailMerge.mailing_label_merge.header.docx").get_child(aw.NodeType.TABLE, 0, True).get_text().strip())

        self.assertEqual(
            "John\aDoe\a\a",
            aw.Document(ARTIFACTS_DIR + "MailMerge.mailing_label_merge.data.docx").get_child(aw.NodeType.TABLE, 0, True).get_text().strip())

        doc = aw.Document(ARTIFACTS_DIR + "MailMerge.mailing_label_merge.docx")

        self.assertEqual(2, doc.range.fields.count)

        settings = doc.mail_merge_settings

        self.assertEqual(ARTIFACTS_DIR + "MailMerge.mailing_label_merge.header.docx", settings.header_source)
        self.assertEqual(ARTIFACTS_DIR + "MailMerge.mailing_label_merge.data.docx", settings.data_source)
        self.assertEqual("SELECT * FROM " + settings.data_source, settings.query)
        self.assertEqual(aw.settings.MailMergeMainDocumentType.MAILING_LABELS, settings.main_document_type)
        self.assertEqual(aw.settings.MailMergeDataType.TEXT_FILE, settings.data_type)
        self.assertTrue(settings.link_to_query)
        self.assertTrue(settings.view_merged_data)

    def test_odso_field_map_data_collection(self):

        #ExStart
        #ExFor:Odso.field_map_datas
        #ExFor:OdsoFieldMapData
        #ExFor:OdsoFieldMapData.clone
        #ExFor:OdsoFieldMapData.column
        #ExFor:OdsoFieldMapData.mapped_name
        #ExFor:OdsoFieldMapData.name
        #ExFor:OdsoFieldMapData.type
        #ExFor:OdsoFieldMapDataCollection
        #ExFor:OdsoFieldMapDataCollection.add(OdsoFieldMapData)
        #ExFor:OdsoFieldMapDataCollection.clear
        #ExFor:OdsoFieldMapDataCollection.count
        #ExFor:OdsoFieldMapDataCollection.__iter__
        #ExFor:OdsoFieldMapDataCollection.__getitem__(int)
        #ExFor:OdsoFieldMapDataCollection.remove_at(int)
        #ExFor:OdsoFieldMappingType
        #ExSummary:Shows how to access the collection of data that maps data source columns to merge fields.
        doc = aw.Document(MY_DIR + "Odso data.docx")

        # This collection defines how a mail merge will map columns from a data source
        # to predefined MERGEFIELD, ADDRESSBLOCK and GREETINGLINE fields.
        data_collection = doc.mail_merge_settings.odso.field_map_datas
        self.assertEqual(30, data_collection.count)

        for index, data in enumerate(data_collection):
            print(f"Field map data index {index}, type \"{data.type}\":")

            if data.type != aw.settings.OdsoFieldMappingType.NULL:
                print(f"\tColumn \"{data.name}\", number {data.column} mapped to merge field \"{data.mapped_name}\".")
            else:
                print("\tNo valid column to field mapping data present.")

        # Clone the elements in this collection.
        #Assert.are_not_equal(data_collection[0], data_collection[0].clone())

        # Use the "remove_at" method elements individually by index.
        data_collection.remove_at(0)

        self.assertEqual(29, data_collection.count)

        # Use the "clear" method to clear the entire collection at once.
        data_collection.clear()

        self.assertEqual(0, data_collection.count)
        #ExEnd

    def test_odso_recipient_data_collection(self):

        #ExStart
        #ExFor:Odso.recipient_datas
        #ExFor:OdsoRecipientData
        #ExFor:OdsoRecipientData.active
        #ExFor:OdsoRecipientData.clone
        #ExFor:OdsoRecipientData.column
        #ExFor:OdsoRecipientData.hash
        #ExFor:OdsoRecipientData.unique_tag
        #ExFor:OdsoRecipientDataCollection
        #ExFor:OdsoRecipientDataCollection.add(OdsoRecipientData)
        #ExFor:OdsoRecipientDataCollection.clear
        #ExFor:OdsoRecipientDataCollection.count
        #ExFor:OdsoRecipientDataCollection.__iter__
        #ExFor:OdsoRecipientDataCollection.__getitem__(int)
        #ExFor:OdsoRecipientDataCollection.remove_at(int)
        #ExSummary:Shows how to access the collection of data that designates which merge data source records a mail merge will exclude.
        doc = aw.Document(MY_DIR + "Odso data.docx")

        data_collection = doc.mail_merge_settings.odso.recipient_datas

        self.assertEqual(70, data_collection.count)

        for index, data in enumerate(data_collection):
            print(f'Odso recipient data index {index} will {"" if data.active else "not "}be imported upon mail merge.')
            print(f"\tColumn #{data.column}")
            print(f"\tHash code: {data.hash}")
            print(f"\tContents array length: {data.unique_tag.length}")

        # We can clone the elements in this collection.
        self.assertNotEqual(data_collection[0], data_collection[0].clone())

        # We can also remove elements individually, or clear the entire collection at once.
        data_collection.remove_at(0)

        self.assertEqual(69, data_collection.count)

        data_collection.clear()

        self.assertEqual(0, data_collection.count)
        #ExEnd

    def test_change_field_update_culture_source(self):

        #ExStart
        #ExFor:Document.field_options
        #ExFor:FieldOptions
        #ExFor:FieldOptions.field_update_culture_source
        #ExFor:FieldUpdateCultureSource
        #ExSummary:Shows how to specify the source of the culture used for date formatting during a field update or mail merge.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert two merge fields with German locale.
        builder.font.locale_id = CultureInfo("de-DE").lcid
        builder.insert_field("MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"")
        builder.write(" - ")
        builder.insert_field("MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"")

        # Set the current culture to US English after preserving its original value in a variable.
        current_culture = Thread.current_thread.current_culture
        Thread.current_thread.current_culture = CultureInfo("en-US")

        # This merge will use the current thread's culture to format the date, US English.
        doc.mail_merge.execute(["Date1"], [datetime(2020, 1, 1)])

        # Configure the next merge to source its culture value from the field code. The value of that culture will be German.
        doc.field_options.field_update_culture_source = aw.fields.FieldUpdateCultureSource.FIELD_CODE
        doc.mail_merge.execute(["Date2"], [datetime(2020, 1, 1)])

        # The first merge result contains a date formatted in English, while the second one is in German.
        self.assertEqual("Wednesday, 1 January 2020 - Mittwoch, 1 Januar 2020", doc.range.text.strip())

        # Restore the thread's original culture.
        Thread.current_thread.current_culture = current_culture
        #ExEnd

    def test_restart_lists_at_each_section(self):

        #ExStart
        #ExFor:MailMerge.restart_lists_at_each_section
        #ExSummary:Shows how to control whether or not list numbering is restarted at each section when mail merge is performed.
        doc = aw.Document(MY_DIR + "Section breaks with numbering.docx")

        doc.mail_merge.restart_lists_at_each_section = False
        doc.mail_merge.execute("", object())

        doc.save(ARTIFACTS_DIR + "MailMerge.restart_lists_at_each_section.pdf")
        #ExEnd

    def test_remove_last_empty_paragraph(self):

        #ExStart
        #ExFor:DocumentBuilder.insert_html(str,HtmlInsertOptions)
        #ExSummary:Shows how to use options while inserting html.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_field(" MERGEFIELD Name ")
        builder.insert_paragraph()
        builder.insert_field(" MERGEFIELD EMAIL ")
        builder.insert_paragraph()

        # By default "DocumentBuilder.insert_html" inserts a HTML fragment that ends with a block-level HTML element,
        # it normally closes that block-level element and inserts a paragraph break.
        # As a result, a new empty paragraph appears after inserted document.
        # If we specify "HtmlInsertOptions.REMOVE_LAST_EMPTY_PARAGRAPH", those extra empty paragraphs will be removed.
        builder.move_to_merge_field("NAME")
        builder.insert_html("<p>John Smith</p>", aw.HtmlInsertOptions.USE_BUILDER_FORMATTING | aw.HtmlInsertOptions.REMOVE_LAST_EMPTY_PARAGRAPH)
        builder.move_to_merge_field("EMAIL")
        builder.insert_html("<p>jsmith@example.com</p>", aw.HtmlInsertOptions.USE_BUILDER_FORMATTING)

        doc.save(ARTIFACTS_DIR + "MailMerge.remove_last_empty_paragraph.docx")
        #ExEnd

        self.assertEqual(4, doc.first_section.body.paragraphs.count)
