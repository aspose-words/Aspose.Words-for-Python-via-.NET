# Copyright (c) 2001-2022 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import io

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, IMAGE_DIR, IMAGE_URL

class ExMailMergeEvent(ApiExampleBase):

    #ExStart
    #ExFor:DocumentBuilder.insert_html(str)
    #ExFor:MailMerge.field_merging_callback
    #ExFor:IFieldMergingCallback
    #ExFor:FieldMergingArgs
    #ExFor:FieldMergingArgsBase
    #ExFor:FieldMergingArgsBase.field
    #ExFor:FieldMergingArgsBase.document_field_name
    #ExFor:FieldMergingArgsBase.document
    #ExFor:IFieldMergingCallback.field_merging
    #ExFor:FieldMergingArgs.text
    #ExSummary:Shows how to execute a mail merge with a custom callback that handles merge data in the form of HTML documents.
    def test_merge_html(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_field(r"MERGEFIELD  html_Title  \b Content")
        builder.insert_field(r"MERGEFIELD  html_Body  \b Content")

        merge_data = [
            "<html>" +
                "<h1>" +
                    "<span style=\"color: #0000ff; font-family: Arial;\">Hello World!</span>" +
                "</h1>" +
            "</html>",

            "<html>" +
                "<blockquote>" +
                    "<p>Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</p>" +
                "</blockquote>" +
            "</html>"
        ]

        doc.mail_merge.field_merging_callback = ExMailMergeEvent.HandleMergeFieldInsertHtml()
        doc.mail_merge.execute(["html_Title", "html_Body"], merge_data)

        doc.save(ARTIFACTS_DIR + "MailMergeEvent.merge_html.docx")

    class HandleMergeFieldInsertHtml(aw.mailmerging.IFieldMergingCallback):
        """If the mail merge encounters a MERGEFIELD whose name starts with the "html_" prefix,
        this callback parses its merge data as HTML content and adds the result to the document location of the MERGEFIELD."""

        def field_merging(self, args: aw.mailmerging.FieldMergingArgs):
            """Called when a mail merge merges data into a MERGEFIELD."""

            if args.document_field_name.startswith("html_") and "\\b" in args.field.get_field_code():

                # Add parsed HTML data to the document's body.
                builder = aw.DocumentBuilder(args.document)
                builder.move_to_merge_field(args.document_field_name)
                builder.insert_html(str(args.field_value))

                # Since we have already inserted the merged content manually,
                # we will not need to respond to this event by returning content via the "text" property.
                args.text = ""

        def image_field_merging(self, args: aw.mailmerging.ImageFieldMergingArgs):

            # Do nothing.
            pass

    #ExEnd

    #ExStart
    #ExFor:FieldMergingArgsBase.field_value
    #ExSummary:Shows how to edit values that MERGEFIELDs receive as a mail merge takes place.
    def test_field_formats(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert some MERGEFIELDs with format switches that will edit the values they will receive during a mail merge.
        builder.insert_field("MERGEFIELD text_Field1 \\* Caps", None)
        builder.write(", ")
        builder.insert_field("MERGEFIELD text_Field2 \\* Upper", None)
        builder.write(", ")
        builder.insert_field("MERGEFIELD numeric_Field1 \\# 0.0", None)

        builder.document.mail_merge.field_merging_callback = ExMailMergeEvent.FieldValueMergingCallback()

        builder.document.mail_merge.execute(
            ["text_Field1", "text_Field2", "numeric_Field1"],
            ["Field 1", "Field 2", 10])

        self.assertEqual("Merge Value For \"Text_Field1\": Field 1, MERGE VALUE FOR \"TEXT_FIELD2\": FIELD 2, 10000.0", doc.get_text().strip())

    class FieldValueMergingCallback(aw.mailmerging.IFieldMergingCallback):
        """Edits the values that MERGEFIELDs receive during a mail merge.
        The name of a MERGEFIELD must have a prefix for this callback to take effect on its value."""

        def field_merging(self, args: aw.mailmerging.FieldMergingArgs):
            """Called when a mail merge merges data into a MERGEFIELD."""

            if args.field_name.startswith("text_"):
                args.field_value = f"Merge value for \"{args.field_name}\": {args.field_value}"
            elif args.field_name.startswith("numeric_"):
                args.field_value = int(args.field_value) * 1000

        def image_field_merging(self, args: aw.mailmerging.ImageFieldMergingArgs):

            # Do nothing.
            pass

    #ExEnd

    #ExStart
    #ExFor:DocumentBuilder.move_to_merge_field(str)
    #ExFor:FieldMergingArgsBase.field_name
    #ExFor:FieldMergingArgsBase.table_name
    #ExFor:FieldMergingArgsBase.record_index
    #ExSummary:Shows how to insert checkbox form fields into MERGEFIELDs as merge data during mail merge.
    def test_insert_check_box(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Use MERGEFIELDs with "TableStart"/"TableEnd" tags to define a mail merge region
        # which belongs to a data source named "StudentCourse" and has a MERGEFIELD which accepts data from a column named "CourseName".
        builder.start_table()
        builder.insert_cell()
        builder.insert_field(" MERGEFIELD  TableStart:StudentCourse ")
        builder.insert_cell()
        builder.insert_field(" MERGEFIELD  CourseName ")
        builder.insert_cell()
        builder.insert_field(" MERGEFIELD  TableEnd:StudentCourse ")
        builder.end_table()

        doc.mail_merge.field_merging_callback = ExMailMergeEvent.HandleMergeFieldInsertCheckBox()

        data_table = ExMailMergeEvent.get_student_course_data_table()

        doc.mail_merge.execute_with_regions(data_table)
        doc.save(ARTIFACTS_DIR + "MailMergeEvent.insert_check_box.docx")
        self.mail_merge_matches_data_table(data_table, aw.Document(ARTIFACTS_DIR + "MailMergeEvent.insert_check_box.docx"), False) #ExSkip

    class HandleMergeFieldInsertCheckBox(aw.mailmerging.IFieldMergingCallback):
        """Upon encountering a MERGEFIELD with a specific name, inserts a check box form field instead of merge data text."""

        def __init__(self):
            self.check_box_count = 0

        def field_merging(self, args: aw.mailmerging.FieldMergingArgs):
            """Called when a mail merge merges data into a MERGEFIELD."""

            if args.document_field_name == "CourseName":
                self.assertEqual("StudentCourse", args.table_name)

                builder = aw.DocumentBuilder(args.document)
                builder.move_to_merge_field(args.field_name)
                builder.insert_check_box(args.document_field_name + self.check_box_count, False, 0)

                field_value = args.field_value.to_string()

                # In this case, for every record index 'n', the corresponding field value is "Course n".
                self.assertEqual(ord(field_value[7]), args.record_index)

                builder.write(field_value)
                self.check_box_count += 1

        def image_field_merging(self, args: aw.mailmerging.ImageFieldMergingArgs):

            # Do nothing.
            pass

    @staticmethod
    def get_student_course_data_table() -> DataTable:
        """Creates a mail merge data source."""

        data_table = DataTable("StudentCourse")
        data_table.columns.add("CourseName")
        for i in range(10):

            datarow = data_table.new_row()
            data_table.rows.add(datarow)
            datarow[0] = "Course " + i

        return data_table

    #ExEnd

    #ExStart
    #ExFor:MailMerge.execute_with_regions(DataTable)
    #ExSummary:Demonstrates how to format cells during a mail merge.
    def test_alternating_rows(self):

        doc = aw.Document(MY_DIR + "Mail merge destination - Northwind suppliers.docx")

        doc.mail_merge.field_merging_callback = ExMailMergeEvent.HandleMergeFieldAlternatingRows()

        data_table = ExMailMergeEvent.get_suppliers_data_table()
        doc.mail_merge.execute_with_regions(data_table)

        doc.save(ARTIFACTS_DIR + "MailMergeEvent.alternating_rows.docx")
        self.mail_merge_matches_data_table(data_table, aw.Document(ARTIFACTS_DIR + "MailMergeEvent.alternating_rows.docx"), False) #ExSkip

    class HandleMergeFieldAlternatingRows(aw.mailmerging.IFieldMergingCallback):
        """Formats table rows as a mail merge takes place to alternate between two colors on odd/even rows."""

        def __init__(self):
            self.builder = None
            self.row_idx = 0

        def field_merging(self, args: aw.mailmerging.FieldMergingArgs):
            """Called when a mail merge merges data into a MERGEFIELD."""

            if self.builder is None:
                self.builder = aw.DocumentBuilder(args.document)

            # This is true of we are on the first column, which means we have moved to a new row.
            if args.field_name == "CompanyName":
                row_color = drawing.Color.from_argb(213, 227, 235) if self.row_idx % 2 == 0 else drawing.Color.from_argb(242, 242, 242)

                for col_idx in range(4):
                    self.builder.move_to_cell(0, self.row_idx, col_idx, 0)
                    self.builder.cell_format.shading.background_pattern_color = row_color

                self.row_idx += 1

        def image_field_merging(self, args: aw.mailmerging.ImageFieldMergingArgs):

            # Do nothing.
            pass

    @staticmethod
    def get_suppliers_data_table() -> DataTable:
        """Creates a mail merge data source."""

        data_table = DataTable("Suppliers")
        data_table.columns.add("CompanyName")
        data_table.columns.add("ContactName")
        for i in range(10):
            datarow = data_table.new_row()
            data_table.rows.add(datarow)
            datarow[0] = "Company " + i
            datarow[1] = "Contact " + i

        return data_table

    #ExEnd

    def test_image_from_url(self):

        #ExStart
        #ExFor:MailMerge.execute(List[str],List[object])
        #ExSummary:Shows how to merge an image from a URI as mail merge data into a MERGEFIELD.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # MERGEFIELDs with "Image:" tags will receive an image during a mail merge.
        # The string after the colon in the "Image:" tag corresponds to a column name
        # in the data source whose cells contain URIs of image files.
        builder.insert_field("MERGEFIELD  Image:logo_FromWeb ")
        builder.insert_field("MERGEFIELD  Image:logo_FromFileSystem ")

        # Create a data source that contains URIs of images that we will merge.
        # A URI can be a web URL that points to an image, or a local file system filename of an image file.
        columns = ["logo_FromWeb", "logo_FromFileSystem"]
        uris = [IMAGE_URL, IMAGE_DIR + "Logo.jpg"]

        # Execute a mail merge on a data source with one row.
        doc.mail_merge.execute(columns, uris)

        doc.save(ARTIFACTS_DIR + "MailMergeEvent.image_from_url.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "MailMergeEvent.image_from_url.docx")

        image_shape = doc.get_child(aw.NodeType.SHAPE, 1, True).as_shape()

        self.verify_image_in_shape(400, 400, aw.drawing.image_type.JPEG, image_shape)

        image_shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()

        self.verify_image_in_shape(252, 213, aw.drawing.image_type.PNG, image_shape)

    #ExStart
    #ExFor:MailMerge.field_merging_callback
    #ExFor:MailMerge.execute_with_regions(IDataReader,str)
    #ExFor:IFieldMergingCallback
    #ExFor:ImageFieldMergingArgs
    #ExFor:IFieldMergingCallback.field_merging
    #ExFor:IFieldMergingCallback.image_field_merging
    #ExFor:ImageFieldMergingArgs.image_stream
    #ExSummary:Shows how to insert images stored in a database BLOB field into a report.
    def test_image_from_blob(self):

        doc = aw.Document(MY_DIR + "Mail merge destination - Northwind employees.docx")

        doc.mail_merge.field_merging_callback = ExMailMergeEvent.HandleMergeImageFieldFromBlob()

        conn_string = f"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={DatabaseDir + 'Northwind.mdb'};"
        query = "SELECT FirstName, LastName, Title, Address, City, Region, Country, PhotoBLOB FROM Employees"

        with OleDbConnection(conn_string) as conn:
            conn.open()

            # Open the data reader, which needs to be in a mode that reads all records at once.
            cmd = OleDbCommand(query, conn)
            data_reader = cmd.execute_reader()

            doc.mail_merge.execute_with_regions(data_reader, "Employees")

        doc.save(ARTIFACTS_DIR + "MailMergeEvent.image_from_blob.docx")
        self.mail_merge_matches_query_result(DatabaseDir + "Northwind.mdb", query, aw.Document(ARTIFACTS_DIR + "MailMergeEvent.image_from_blob.docx"), False) #ExSkip

    class HandleMergeImageFieldFromBlob(aw.mailmerging.IFieldMergingCallback):

        def field_merging(self, args: aw.mailmerging.FieldMergingArgs):

            # Do nothing.
            pass

        def image_field_merging(self, args: aw.mailmerging.ImageFieldMergingArgs):
            """This is called when a mail merge encounters a MERGEFIELD in the document with an "Image:" tag in its name."""

            args.image_stream = io.BytesIO(args.field_value)

    #ExEnd
