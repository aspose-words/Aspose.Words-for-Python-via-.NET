import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR, IMAGES_DIR

class WorkingWithFields(DocsExamplesBase):

    def test_mail_merge_form_fields(self):

        #ExStart:MailMergeFormFields
        #GistId:a66c4234a53ab6f90df96f76cb549ec1
        doc = aw.Document(MY_DIR + "Mail merge destinations - Fax.docx")

        # Setup mail merge event handler to do the custom work.
        doc.mail_merge.field_merging_callback = HandleMergeField()
        # Trim trailing and leading whitespaces mail merge values.
        doc.mail_merge.trim_whitespaces = False

        field_names = ["RecipientName", "SenderName", "FaxNumber", "PhoneNumber",
            "Subject", "Body", "Urgent", "ForReview", "PleaseComment"]

        field_values = ["Josh", "Jenny", "123456789", "", "Hello",
            "<b>HTML Body Test message 1</b>", True, False, True]

        doc.mail_merge.execute(field_names, field_values)

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.mail_merge_form_fields.docx")
        #ExEnd:MailMergeFormFields

    def test_mail_merge_image_field(self):

        #ExStart:MailMergeImageField
        #GistId:7dd46d9612db0a89636536b4b8f2a935
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("{{#foreach example}}")
        builder.writeln("{{Image(126pt;126pt):stempel}}")
        builder.writeln("{{/foreach example}}")

        doc.mail_merge.use_non_merge_fields = True
        doc.mail_merge.trim_whitespaces = True
        doc.mail_merge.use_whole_paragraph_as_region = False
        doc.mail_merge.cleanup_options = (aw.mailmerging.MailMergeCleanupOptions.REMOVE_EMPTY_TABLE_ROWS
            | aw.mailmerging.MailMergeCleanupOptions.REMOVE_CONTAINING_FIELDS
            | aw.mailmerging.MailMergeCleanupOptions.REMOVE_UNUSED_REGIONS
            | aw.mailmerging.MailMergeCleanupOptions.REMOVE_UNUSED_FIELDS)

        doc.mail_merge.field_merging_callback = ImageFieldMergingHandler()
        doc.mail_merge.execute_with_regions(DataSourceRoot())

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.mail_merge_image_field.docx")
        #ExEnd:MailMergeImageField

    def test_handle_mail_merge_switches(self):

        doc = aw.Document(MY_DIR + "Field sample - MERGEFIELD.docx")

        doc.mail_merge.field_merging_callback = MailMergeSwitches()

        html = """<html>
                <h1>Hello world!</h1>
        </html>"""

        doc.mail_merge.execute(["htmlField1"], [html])

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.handle_mail_merge_switches.docx")

    def test_field_next(self):

        #ExStart:FieldNext
        #GistId:81ec38c287f6a1e18368813763a6c7d1
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Use NextIf field. A NEXTIF field has the same function as a NEXT field,
        # but it skips to the next row only if a statement constructed by the following 3 properties is true.
        field_next_if = builder.insert_field(aw.fields.FieldType.FIELD_NEXT_IF, True).as_field_next_if()

        # Or use SkipIf field.
        field_skip_if = builder.insert_field(aw.fields.FieldType.FIELD_SKIP_IF, True).as_field_skip_if()

        field_next_if.left_expression = "5"
        field_next_if.right_expression = "2 + 3"
        field_next_if.comparison_operator = "="

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.field_next.docx")
        #ExEnd:FieldNext

#ExStart:HandleMergeField
#GistId:7dd46d9612db0a89636536b4b8f2a935
class HandleMergeField(aw.mailmerging.IFieldMergingCallback):
    """This handler is called for every mail merge field found in the document,
    for every record found in the data source."""

    def __init__(self):
        self.builder = None

    def field_merging(self, args: aw.mailmerging.FieldMergingArgs):
        if self.builder is None:
            self.builder = aw.DocumentBuilder(args.document)

        # We decided that we want all boolean values to be output as check box form fields.
        if isinstance(args.field_value, bool):
            # Move the "cursor" to the current merge field.
            self.builder.move_to_merge_field(args.field_name)

            check_box_name = f"{args.field_name}{args.record_index}"

            self.builder.insert_check_box(check_box_name, args.field_value, 0)

            return

        if args.field_name == "Body":
            self.builder.move_to_merge_field(args.field_name)
            self.builder.insert_html(args.field_value)

        elif args.field_name == "Subject":
            self.builder.move_to_merge_field(args.field_name)
            text_input_name = f"{args.field_name}{args.record_index}"
            self.builder.insert_text_input(text_input_name, aw.fields.TextFormFieldType.REGULAR, "", args.field_value, 0)

    #ExStart:ImageFieldMerging
    #GistId:a66c4234a53ab6f90df96f76cb549ec1
    def image_field_merging(self, args: aw.mailmerging.ImageFieldMergingArgs):
        args.image_file_name = IMAGES_DIR + "Logo.jpg"
        args.image_width.value = 200
        args.image_height = aw.fields.MergeFieldImageDimension(200, aw.fields.MergeFieldImageDimensionUnit.PERCENT)
    #ExEnd:ImageFieldMerging
#ExEnd:HandleMergeField

#ExStart:ImageFieldMergingHandler
#GistId:7dd46d9612db0a89636536b4b8f2a935
class ImageFieldMergingHandler(aw.mailmerging.IFieldMergingCallback):

    def field_merging(self, args: aw.mailmerging.FieldMergingArgs):
        # Implementation is not required.
        pass

    def image_field_merging(self, args: aw.mailmerging.ImageFieldMergingArgs):
        shape = aw.drawing.Shape(args.document, aw.drawing.ShapeType.IMAGE)
        shape.width = 126
        shape.height = 126
        shape.wrap_type = aw.drawing.WrapType.SQUARE

        shape.image_data.set_image(IMAGES_DIR + "Logo.jpg")

        args.shape = shape
#ExEnd:ImageFieldMergingHandler

#ExStart:DataSourceRoot
#GistId:7dd46d9612db0a89636536b4b8f2a935
class DataSourceRoot(aw.mailmerging.IMailMergeDataSourceRoot):

    def get_data_source(self, table_name: str):
        return DataSourceRoot.DataSource()

    class DataSource(aw.mailmerging.IMailMergeDataSource):

        def __init__(self):
            self.next = True

        @property
        def table_name(self):
            return "example"

        def move_next(self):
            result = self.next
            self.next = False
            return result

        def get_child_data_source(self, table_name: str):
            return None

        def get_value(self, field_name: str, field_value):
            return False
#ExEnd:DataSourceRoot

#ExStart:HandleMailMergeSwitches
class MailMergeSwitches(aw.mailmerging.IFieldMergingCallback):

    def field_merging(self, args: aw.mailmerging.FieldMergingArgs):
        if args.field_name.upper().startswith("HTML"):
            if args.field.get_field_code().find("\\b") >= 0:
                field = args.field

                builder = aw.DocumentBuilder(args.document)
                builder.move_to_merge_field(args.document_field_name, True, False)
                builder.write(field.text_before)
                builder.insert_html(args.field_value)

                args.text = ""

    def image_field_merging(self, args: aw.mailmerging.ImageFieldMergingArgs):
        pass
#ExEnd:HandleMailMergeSwitches
