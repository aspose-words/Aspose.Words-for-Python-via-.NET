# Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
import unittest
from typing import Dict, List

import aspose.words as aw

from api_example_base import ApiExampleBase, ARTIFACTS_DIR

@unittest.skip("type 'aspose.words.mailmerging.IMailMergeDataSource' is not an acceptable base type ")
class ExMailMergeCustom(ApiExampleBase):

    #ExStart
    #ExFor:IMailMergeDataSource
    #ExFor:IMailMergeDataSource.table_name
    #ExFor:IMailMergeDataSource.move_next
    #ExFor:IMailMergeDataSource.get_value
    #ExFor:IMailMergeDataSource.get_child_data_source
    #ExFor:MailMerge.execute(IMailMergeDataSourceCore)
    #ExSummary:Shows how to execute a mail merge with a data source in the form of a custom object.
    def test_custom_data_source(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.insert_field(" MERGEFIELD FullName ")
        builder.insert_paragraph()
        builder.insert_field(" MERGEFIELD Address ")

        customers = [] # type: List[Customer]
        customers.add(Customer("Thomas Hardy", "120 Hanover Sq., London"))
        customers.add(Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"))

        # To use a custom object as a data source, it must implement the IMailMergeDataSource interface.
        data_source = ExMailMergeCustom.CustomerMailMergeDataSource(customers)

        doc.mail_merge.execute(data_source)

        doc.save(ARTIFACTS_DIR + "MailMergeCustom.custom_data_source.docx")
        self._test_custom_data_source(customers, aw.Document(ARTIFACTS_DIR + "MailMergeCustom.custom_data_source.docx")) #ExSkip

    class Customer:
        """An example of a "data entity" class in your application."""

        def __init__(self, full_name: str, address: str):

            self.full_name = full_name
            self.address = address

    class CustomerMailMergeDataSource(aw.mailmerging.IMailMergeDataSource):
        """A custom mail merge data source that you implement to allow Aspose.Words
        to mail merge data from your Customer objects into Microsoft Word documents."""

        def __init__(self, customers: 'List[ExMailMergeCustom.Customer]'):

            self.customers = customers

            # When we initialize the data source, its position must be before the first record.
            self.record_index = -1

        @property
        def table_name(self) -> str:
            """The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions."""

            return "Customer"

        def get_value(self, field_name: str) -> str:
            """Aspose.Words calls this method to get a value for every data field."""

            if field_name == "FullName":
                return self.customers[self.record_index].full_name

            if field_name ==  "Address":
                return self.customers[self.record_index].address

            # Return "None" to the Aspose.Words mail merge engine to signify
            # that we could not find a field with this name.
            return None

        def move_next(self) -> bool:
            """A standard implementation for moving to a next record in a collection."""

            if not self.is_eof:
                self.record_index += 1

            return not self.is_eof

        def get_child_data_source(self, table_name: str):

            return None

        @property
        def is_eof(self) -> bool:

            return self.record_index >= len(self.customers)

    #ExEnd

    def _test_custom_data_source(self, customer_list: List[ExMailMergeCustom.Customer], doc: aw.Document):

        merge_data = []

        for customer in customer_list:
            merge_data.append([customer.full_name, customer.address])

        self.mail_merge_matches_array(merge_data, doc, True)

    #ExStart
    #ExFor:IMailMergeDataSourceRoot
    #ExFor:IMailMergeDataSourceRoot.get_data_source(str)
    #ExFor:MailMerge.execute_with_regions(IMailMergeDataSourceRoot)
    #ExSummary:Performs mail merge from a custom data source with master-detail data.
    def test_custom_data_source_root(self):

        # Create a document with two mail merge regions named "Washington" and "Seattle".
        mail_merge_regions = ["Vancouver", "Seattle"]
        doc = ExMailMergeCustom.create_source_document_with_mail_merge_regions(mail_merge_regions)

        # Create two data sources for the mail merge.
        employees_washington_branch = [] # type: List[Employee]
        employees_washington_branch.add(ExMailMergeCustom.Employee("John Doe", "Sales"))
        employees_washington_branch.add(ExMailMergeCustom.Employee("Jane Doe", "Management"))

        employees_seattle_branch = [] # type: List[Employee]
        employees_seattle_branch.add(ExMailMergeCustom.Employee("John Cardholder", "Management"))
        employees_seattle_branch.add(ExMailMergeCustom.Employee("Joe Bloggs", "Sales"))

        # Register our data sources by name in a data source root.
        # If we are about to use this data source root in a mail merge with regions,
        # each source's registered name must match the name of an existing mail merge region in the mail merge source document.
        source_root = ExMailMergeCustom.DataSourceRoot()
        source_root.register_source(mail_merge_regions[0], ExMailMergeCustom.EmployeeListMailMergeSource(employees_washington_branch))
        source_root.register_source(mail_merge_regions[1], ExMailMergeCustom.EmployeeListMailMergeSource(employees_seattle_branch))

        # Since we have consecutive mail merge regions, we would normally have to perform two mail merges.
        # However, one mail merge source with a data root can fill in multiple regions
        # if the root contains tables with corresponding names/column names.
        doc.mail_merge.execute_with_regions(source_root)

        doc.save(ARTIFACTS_DIR + "MailMergeCustom.custom_data_source_root.docx")
        self._test_custom_data_source_root(mail_merge_regions, source_root, aw.Document(ARTIFACTS_DIR + "MailMergeCustom.custom_data_source_root.docx")) #ExSkip

    @staticmethod
    def create_source_document_with_mail_merge_regions(regions: List[str]) -> aw.Document:
        """Create a document that contains consecutive mail merge regions, with names designated by the input array,
        for a data table of employees."""

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        for region in regions:
            builder.writeln("\n" + region + " branch: ")
            builder.insert_field(" MERGEFIELD TableStart:" + region)
            builder.insert_field(" MERGEFIELD FullName")
            builder.write(", ")
            builder.insert_field(" MERGEFIELD Department")
            builder.insert_field(" MERGEFIELD TableEnd:" + region)

        return doc

    class Employee:
        """An example of a "data entity" class in your application."""

        def __init__(self, full_name: str, department: str):

            self.full_name = full_name
            self.department = department

    class DataSourceRoot(aw.mailmerging.IMailMergeDataSourceRoot):
        """Data source root that can be passed directly into a mail merge which can register and contain many child data sources.
        These sources must all implement IMailMergeDataSource, and are registered and differentiated by a name
        which corresponds to a mail merge region that will read the respective data."""

        def __init__(self):
            self.sources: Dict[str, ExMailMergeCustom.EmployeeListMailMergeSource] = {}

        def get_data_source(self, table_name: str) -> aw.mailmerging.IMailMergeDataSource:

            source = self.sources[table_name]
            source.reset()
            return self.sources[table_name]

        def register_source(self, source_name: str, source: ExMailMergeCustom.EmployeeListMailMergeSource):

            self.sources.add(source_name, source)

    class EmployeeListMailMergeSource(aw.mailmerging.IMailMergeDataSource):
        """Custom mail merge data source."""

        def __init__(self, employees: List[ExMailMergeCustom.Employee]):

            self.employees = employees
            self.record_index = -1

        def move_next(self) -> bool:
            """A standard implementation for moving to a next record in a collection."""

            if not self.is_eof:
                self.record_index += 1

            return not self.is_eof

        @property
        def is_eof(self) -> bool:

            return self.record_index >= len(self.employees)

        def test_reset(self):

            self.record_index = -1

        @property
        def table_name(self) -> str:
            """The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions."""

            return "Employees"

        def get_value(self, field_name: str) -> str:
            """Aspose.Words calls this method to get a value for every data field."""

            if field_name == "FullName":
                return self.employees[self.record_index].full_name

            if field_name == "Department":
                return self.employees[self.record_index].department

            # Return "None" to the Aspose.Words mail merge engine to signify
            # that we could not find a field with this name.
            return None

        def get_child_data_source(self, table_name: str) -> aw.mailmerging.IMailMergeDataSource:
            """Child data sources are for nested mail merges."""

            raise NotImplementedError()

    #ExEnd

    def _test_custom_data_source_root(self, registered_sources: List[str], source_root: ExMailMergeCustom.DataSourceRoot, doc: aw.Document):

        data_table = DataTable()
        data_table.columns.add("FullName")
        data_table.columns.add("Department")

        for source_name in registered_sources:

            source = source_root.get_data_source(source_name)
            while source.move_next():
                full_name = source.get_value("FullName")
                department = source.get_value("Department")

                data_table.rows.add([full_name, department])

        self.mail_merge_matches_data_table(data_table, doc, False)
