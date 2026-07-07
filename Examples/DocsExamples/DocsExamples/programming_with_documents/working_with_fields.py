from datetime import date, datetime
import locale
import re
import sys
import unittest
import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class WorkingWithFields(DocsExamplesBase):

    def test_field_code(self):

        #ExStart:FieldCode
        #GistId:60273b5318979680cf1bdc7bdbebac25
        doc = aw.Document(MY_DIR + "Hyperlinks.docx")

        for field in doc.range.fields:
            field_code = field.get_field_code()
            field_result = field.result
        #ExEnd:FieldCode

    def test_change_field_update_culture_source(self):

        #ExStart:ChangeFieldUpdateCultureSource
        #GistId:0583e2cf3a8d0d6d0ba63e8efc7f2ebf
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert content with German locale.
        builder.font.locale_id = 1031
        builder.insert_field('MERGEFIELD Date1 \\@ "dddd, d MMMM yyyy"')
        builder.write(" - ")
        builder.insert_field('MERGEFIELD Date2 \\@ "dddd, d MMMM yyyy"')

        # Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from
        # set the culture used during field update to the culture used by the field.
        doc.field_options.field_update_culture_source = aw.fields.FieldUpdateCultureSource.FIELD_CODE
        doc.mail_merge.execute(["Date2"], [date(2011, 1, 1)])

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.change_field_update_culture_source.docx")
        #ExEnd:ChangeFieldUpdateCultureSource

    def test_specify_locale_at_field_level(self):

        #ExStart:SpecifyLocaleAtFieldLevel
        #GistId:3432f1bcba1f8c7ab8ea6d7c0badfdf2
        builder = aw.DocumentBuilder()

        field = builder.insert_field(aw.fields.FieldType.FIELD_DATE, True)
        field.locale_id = 1049

        builder.document.save(ARTIFACTS_DIR + "WorkingWithFields.specifylocale_at_fieldlevel.docx")
        #ExEnd:SpecifyLocaleAtFieldLevel

    def test_replace_hyperlinks(self):

        #ExStart:ReplaceHyperlinks
        #GistId:0a77287e9106956f00f83347b104d40b
        doc = aw.Document(MY_DIR + "Hyperlinks.docx")

        for field in doc.range.fields:
            if field.type == aw.fields.FieldType.FIELD_HYPERLINK:
                hyperlink = field.as_field_hyperlink()

                # Some hyperlinks can be local (links to bookmarks inside the document), ignore these.
                if hyperlink.sub_address is not None:
                    continue

                hyperlink.address = "http://www.aspose.com"
                hyperlink.result = "Aspose - The .net & Java Component Publisher"

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.replace_hyperlinks.docx")
        #ExEnd:ReplaceHyperlinks

    def test_rename_merge_fields(self):

        #ExStart:RenameMergeFields
        #GistId:1ead3cc2e51140806a4919331c87eb2d
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_field("MERGEFIELD MyMergeField1 \\* MERGEFORMAT")
        builder.insert_field("MERGEFIELD MyMergeField2 \\* MERGEFORMAT")

        for f in doc.range.fields:
            if f.type == aw.fields.FieldType.FIELD_MERGE_FIELD:
                merge_field = f.as_field_merge_field()
                merge_field.field_name = merge_field.field_name + "_Renamed"
                merge_field.update()

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.rename_merge_fields.docx")
        #ExEnd:RenameMergeFields    

    def test_remove_field(self):

        #ExStart:RemoveField
        #GistId:157a60b5ef694580567e1020eb19c545
        doc = aw.Document(MY_DIR + "Various fields.docx")

        field = doc.range.fields[0]
        field.remove()
        #ExEnd:RemoveField

    def test_unlink_fields(self):

        #ExStart:UnlinkFields
        #GistId:e36d72df7a15934d02a89f614188a967
        doc = aw.Document(MY_DIR + "Various fields.docx")
        doc.unlink_fields()
        #ExEnd:UnlinkFields

    def test_insert_toa_field_without_document_builder(self):

        #ExStart:InsertToaFieldWithoutDocumentBuilder
        #GistId:3432f1bcba1f8c7ab8ea6d7c0badfdf2
        doc = aw.Document()
        para = aw.Paragraph(doc)

        # We want to insert TA and TOA fields like this:
        #  { TA  \c 1 \l "Value 0" }
        #  { TOA  \c 1 }

        field_ta = para.append_field(aw.fields.FieldType.FIELD_TOA_ENTRY, False).as_field_ta()
        field_ta.entry_category = "1"
        field_ta.long_citation = "Value 0"

        doc.first_section.body.append_child(para)

        para = aw.Paragraph(doc)

        field_toa = para.append_field(aw.fields.FieldType.FIELD_TOA, False).as_field_toa()
        field_toa.entry_category = "1"
        doc.first_section.body.append_child(para)

        field_toa.update()

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.insert_toa_field_without_document_builder.docx")
        #ExEnd:InsertToaFieldWithoutDocumentBuilder

    def test_insert_nested_fields(self):

        #ExStart:InsertNestedFields
        #GistId:3432f1bcba1f8c7ab8ea6d7c0badfdf2
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        for _ in range(5):
            builder.insert_break(aw.BreakType.PAGE_BREAK)

        builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)

        # We want to insert a field like this:
        #  IF PAGE <> NUMPAGES "See Next Page" "Last Page"
        field = builder.insert_field("IF ")
        builder.move_to(field.separator)
        builder.insert_field("PAGE")
        builder.write(" <> ")
        builder.insert_field("NUMPAGES")
        builder.write(' "See Next Page" "Last Page" ')

        field.update()

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.insert_nested_fields.docx")
        #ExEnd:InsertNestedFields

    def test_insert_merge_field_using_dom(self):

        #ExStart:InsertMergeFieldUsingDom
        #GistId:3432f1bcba1f8c7ab8ea6d7c0badfdf2
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        para = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)[0].as_paragraph()

        builder.move_to(para)

        # We want to insert a merge field like this:
        #  { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }

        field = builder.insert_field(aw.fields.FieldType.FIELD_MERGE_FIELD, False).as_field_merge_field()

        #  { " MERGEFIELD Test1" }
        field.field_name = "Test1"

        #  { " MERGEFIELD Test1 \\b Test2" }
        field.text_before = "Test2"

        #  { " MERGEFIELD Test1 \\b Test2 \\f Test3 }
        field.text_after = "Test3"

        #  { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m" }
        field.is_mapped = True

        #  { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }
        field.is_vertical_formatting = True

        # Finally update this merge field
        field.update()

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.insert_merge_field_using_dom.docx")
        #ExEnd:InsertMergeFieldUsingDom

    def test_insert_address_block_field_using_dom(self):

        #ExStart:InsertAddressBlockFieldUsingDom
        #GistId:3432f1bcba1f8c7ab8ea6d7c0badfdf2
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        para = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)[0].as_paragraph()

        builder.move_to(para)

        # We want to insert a mail merge address block like this:
        #  { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\" }

        field = builder.insert_field(aw.fields.FieldType.FIELD_ADDRESS_BLOCK, False).as_field_address_block()

        #  { ADDRESSBLOCK \\c 1" }
        field.include_country_or_region_name = "1"

        #  { ADDRESSBLOCK \\c 1 \\d" }
        field.format_address_on_country_or_region = True

        #  { ADDRESSBLOCK \\c 1 \\d \\e Test2 }
        field.excluded_country_or_region_name = "Test2"

        #  { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 }
        field.name_and_address_format = "Test3"

        #  { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\" }
        field.language_id = "Test 4"

        field.update()

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.insert_address_block_field_using_dom.docx")
        #ExEnd:InsertAddressBlockFieldUsingDom

    def test_insert_field_include_text_without_document_builder(self):

        #ExStart:InsertFieldIncludeTextWithoutDocumentBuilder
        #GistId:3432f1bcba1f8c7ab8ea6d7c0badfdf2
        doc = aw.Document()

        para = aw.Paragraph(doc)

        # We want to insert an INCLUDETEXT field like this:
        #  { INCLUDETEXT  "file path" }

        field_include_text = para.append_field(aw.fields.FieldType.FIELD_INCLUDE_TEXT, False).as_field_include_text()
        field_include_text.bookmark_name = "bookmark"
        field_include_text.source_full_name = MY_DIR + "IncludeText.docx"

        doc.first_section.body.append_child(para)

        field_include_text.update()

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.insert_include_field_without_document_builder.docx")
        #ExEnd:InsertFieldIncludeTextWithoutDocumentBuilder

    def test_insert_field_none(self):

        #ExStart:InsertFieldNone
        #GistId:3432f1bcba1f8c7ab8ea6d7c0badfdf2
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        field = builder.insert_field(aw.fields.FieldType.FIELD_NONE, False)

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.insert_field_none.docx")
        #ExEnd:InsertFieldNone

    def test_insert_field(self):

        #ExStart:InsertField
        #GistId:3432f1bcba1f8c7ab8ea6d7c0badfdf2
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_field(r"MERGEFIELD MyFieldName \* MERGEFORMAT")

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.insert_field.docx")
        #ExEnd:InsertField

    def test_insert_field_using_field_builder(self):

        #ExStart:InsertFieldUsingFieldBuilder
        #GistId:3432f1bcba1f8c7ab8ea6d7c0badfdf2
        doc = aw.Document()

        # Prepare IF field with two nested MERGEFIELD fields: { IF "left expression" = "right expression" "Firstname: { MERGEFIELD firstname }" "Lastname: { MERGEFIELD lastname }"}
        field_builder = aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_IF)
        field_builder.add_argument("left expression")
        field_builder.add_argument("=")
        field_builder.add_argument("right expression")
        field_builder.add_argument(
            aw.fields.FieldArgumentBuilder()
                .add_text("Firstname: ")
                .add_field(aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_MERGE_FIELD).add_argument("firstname")))
        field_builder.add_argument(
            aw.fields.FieldArgumentBuilder()
                .add_text("Lastname: ")
                .add_field(aw.fields.FieldBuilder(aw.fields.FieldType.FIELD_MERGE_FIELD).add_argument("lastname")))

        # Insert IF field in exact location            
        field = field_builder.build_and_insert(doc.first_section.body.first_paragraph)
        field.update()

        doc.save(ARTIFACTS_DIR + "Field.insert_field_using_field_builder.docx")
        #ExEnd:InsertFieldUsingFieldBuilder

    def test_insert_author_field(self):

        #ExStart:InsertAuthorField
        #GistId:3432f1bcba1f8c7ab8ea6d7c0badfdf2
        doc = aw.Document()

        para = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)[0].as_paragraph()

        # We want to insert an AUTHOR field like this:
        #  { AUTHOR Test1 }

        field = para.append_field(aw.fields.FieldType.FIELD_AUTHOR, False).as_field_author()
        field.author_name = "Test1" # { AUTHOR Test1 }

        field.update()

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.insert_author_field.docx")
        #ExEnd:InsertAuthorField

    def test_insert_ask_field_with_out_document_builder(self):

        #ExStart:InsertAskFieldWithoutDocumentBuilder
        #GistId:3432f1bcba1f8c7ab8ea6d7c0badfdf2
        doc = aw.Document()

        para = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)[0].as_paragraph()

        # We want to insert an Ask field like this:
        #  { ASK \"Test 1\" Test2 \\d Test3 \\o }

        field = para.append_field(aw.fields.FieldType.FIELD_ASK, False).as_field_ask()

        #  { ASK \"Test 1\" " }
        field.bookmark_name = "Test 1"

        #  { ASK \"Test 1\" Test2 }
        field.prompt_text = "Test2"

        #  { ASK \"Test 1\" Test2 \\d Test3 }
        field.default_response = "Test3"

        #  { ASK \"Test 1\" Test2 \\d Test3 \\o }
        field.prompt_once_on_mail_merge = True

        field.update()

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.insert_ask_field_with_out_document_builder.docx")
        #ExEnd:InsertAskFieldWithoutDocumentBuilder

    def test_insert_advance_field_with_out_document_builder(self):

        #ExStart:InsertAdvanceFieldWithoutDocumentBuilder
        #GistId:3432f1bcba1f8c7ab8ea6d7c0badfdf2
        doc = aw.Document()

        para = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)[0].as_paragraph()

        # We want to insert an Advance field like this:
        #  { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100 }

        field = para.append_field(aw.fields.FieldType.FIELD_ADVANCE, False).as_field_advance()

        #  { ADVANCE \\d 10 " }
        field.down_offset = "10"

        #  { ADVANCE \\d 10 \\l 10 }
        field.left_offset = "10"

        #  { ADVANCE \\d 10 \\l 10 \\r -3.3 }
        field.right_offset = "-3.3"

        #  { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 }
        field.up_offset = "0"

        #  { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 }
        field.horizontal_position = "100"

        #  { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100 }
        field.vertical_position = "100"

        field.update()

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.insert_advance_field_with_out_document_builder.docx")
        #ExEnd:InsertAdvanceFieldWithoutDocumentBuilder

    def test_get_mail_merge_field_names(self):

        #ExStart:GetFieldNames
        #GistId:81ec38c287f6a1e18368813763a6c7d1
        doc = aw.Document()

        field_names = doc.mail_merge.get_field_names()
        #ExEnd:GetFieldNames
        print("\nDocument have {} fields.".format(len(field_names)))

    def test_mapped_data_fields(self):

        #ExStart:MappedDataFields
        #GistId:81ec38c287f6a1e18368813763a6c7d1
        doc = aw.Document()

        doc.mail_merge.mapped_data_fields.add("MyFieldName_InDocument", "MyFieldName_InDataSource")
        #ExEnd:MappedDataFields

    def test_delete_fields(self):

        #ExStart:DeleteFields
        #GistId:d7e717546d32e7dd25f0198a15f8ebf2
        doc = aw.Document()

        doc.mail_merge.delete_fields()
        #ExEnd:DeleteFields

    def test_field_display_results(self):

        #ExStart:FieldDisplayResults
        #GistId:1ead3cc2e51140806a4919331c87eb2d
        #ExStart:UpdateDocFields
        #GistId:365214c8b2e8c065166447871a1499aa
        document = aw.Document(MY_DIR + "Various fields.docx")

        document.update_fields()
        #ExEnd:UpdateDocFields

        for field in document.range.fields:
            print(field.display_result)
        #ExEnd:FieldDisplayResults

    def test_evaluate_if_condition(self):

        #ExStart:EvaluateIfCondition
        #GistId:564a36aefb90914eb4e31faa93f7bed2
        builder = aw.DocumentBuilder()

        field = builder.insert_field("IF 1 = 1", None).as_field_if()
        actual_result = field.evaluate_condition()

        print(actual_result)
        #ExEnd:EvaluateIfCondition

    def test_unlink_fields_in_paragraph(self):

        #ExStart:UnlinkFieldsInParagraph
        #GistId:e36d72df7a15934d02a89f614188a967
        doc = aw.Document(MY_DIR + "Linked fields.docx")

        # Pass the appropriate parameters to convert all IF fields to text that are encountered only in the last
        # paragraph of the document.
        for field in doc.first_section.body.last_paragraph.range.fields:
            if field.type == aw.fields.FieldType.FIELD_IF:
                field.unlink()

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.unlink_fields_in_paragraph.docx")
        #ExEnd:UnlinkFieldsInParagraph

    def test_unlink_fields_in_document(self):

        #ExStart:UnlinkFieldsInDocument
        #GistId:e36d72df7a15934d02a89f614188a967
        doc = aw.Document(MY_DIR + "Linked fields.docx")

        # Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to text.
        for field in doc.range.fields:
            if field.type == aw.fields.FieldType.FIELD_IF:
                field.unlink()

        # Save the document with fields transformed to disk
        doc.save(ARTIFACTS_DIR + "WorkingWithFields.unlink_fields_in_document.docx")
        #ExEnd:UnlinkFieldsInDocument

    def test_unlink_fields_in_body(self):

        #ExStart:UnlinkFieldsInBody
        #GistId:e36d72df7a15934d02a89f614188a967
        doc = aw.Document(MY_DIR + "Linked fields.docx")

        # Pass the appropriate parameters to convert PAGE fields encountered to text only in the body of the first section.
        for field in doc.first_section.body.range.fields:
            if field.type == aw.fields.FieldType.FIELD_PAGE:
                field.unlink()

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.unlink_fields_in_body.docx")
        #ExEnd:UnlinkFieldsInBody

    @unittest.skipUnless(sys.platform.startswith("win"), "requires Windows")
    def test_change_locale(self):

        #ExStart:ChangeLocale
        #GistId:0583e2cf3a8d0d6d0ba63e8efc7f2ebf
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_field("MERGEFIELD Date")

        # Store the current culture so it can be set back once mail merge is complete.
        loc = locale.getlocale()
        # Set to German language so dates and numbers are formatted using this culture during mail merge.
        locale.setlocale(locale.LC_ALL, 'de_DE')

        doc.mail_merge.execute(["Date"], [datetime.now()] )

        locale.setlocale(locale.LC_ALL, loc)

        doc.save(ARTIFACTS_DIR + "WorkingWithFields.change_locale.docx")
        #ExEnd:ChangeLocale
    
    #ExStart:ConvertFieldsToStaticText
    #GistId:e36d72df7a15934d02a89f614188a967
    @staticmethod
    def convert_fields_to_static_text(composite_node, target_field_type):
        """Converts any fields of the specified type found in the descendants of the node into static text.

        :param composite_node: The node in which all descendants of the specified FieldType will be converted to static text.
        :param target_field_type: The FieldType of the field to convert to static text.
        """
        for field in [f for f in composite_node.range.fields if f.type == target_field_type]:
            field.unlink()
    #ExEnd:ConvertFieldsToStaticText
