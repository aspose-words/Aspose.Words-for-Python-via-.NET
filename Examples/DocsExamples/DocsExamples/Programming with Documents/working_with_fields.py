import unittest
import os
import sys
from datetime import date, datetime
import re
import locale

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithFields(docs_base.DocsExamplesBase):

    def test_change_field_update_culture_source(self) :

        #ExStart:ChangeFieldUpdateCultureSource
        #ExStart:DocumentBuilderInsertField
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert content with German locale.
        builder.font.locale_id = 1031
        builder.insert_field("MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"")
        builder.write(" - ")
        builder.insert_field("MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"")
        #ExEnd:DocumentBuilderInsertField

        # Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from
        # set the culture used during field update to the culture used by the field.
        doc.field_options.field_update_culture_source = aw.fields.FieldUpdateCultureSource.FIELD_CODE
        doc.mail_merge.execute(["Date2"], [date(2011, 1, 1)])

        doc.save(docs_base.artifacts_dir + "WorkingWithFields.change_field_update_culture_source.docx")
        #ExEnd:ChangeFieldUpdateCultureSource


    def test_specify_locale_at_field_level(self) :

        #ExStart:SpecifylocaleAtFieldlevel
        builder = aw.DocumentBuilder()

        field = builder.insert_field(aw.fields.FieldType.FIELD_DATE, True)
        field.locale_id = 1049

        builder.document.save(docs_base.artifacts_dir + "WorkingWithFields.specifylocale_at_fieldlevel.docx")
        #ExEnd:SpecifylocaleAtFieldlevel

    def test_replace_hyperlinks(self) :

        #ExStart:ReplaceHyperlinks
        doc = aw.Document(docs_base.my_dir + "Hyperlinks.docx")

        for field in doc.range.fields :

            if field.type == aw.fields.FieldType.FIELD_HYPERLINK:

                hyperlink = field.as_field_hyperlink()

                # Some hyperlinks can be local (links to bookmarks inside the document), ignore these.
                if hyperlink.sub_address != None :
                    continue

                hyperlink.address = "http:#www.aspose.com"
                hyperlink.result = "Aspose - The .net & Java Component Publisher"

        doc.save(docs_base.artifacts_dir + "WorkingWithFields.replace_hyperlinks.docx")
        #ExEnd:ReplaceHyperlinks


    def test_rename_merge_fields(self) :

        #ExStart:RenameMergeFields
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_field("MERGEFIELD MyMergeField1 \\* MERGEFORMAT")
        builder.insert_field("MERGEFIELD MyMergeField2 \\* MERGEFORMAT")

        # Select all field start nodes so we can find the merge fields.
        field_starts = doc.get_child_nodes(aw.NodeType.FIELD_START, True)
        for start in field_starts :

            field_start = start.as_field_start()
            if field_start.field_type == aw.fields.FieldType.FIELD_MERGE_FIELD :

                merge_field = self.MergeField(field_start)
                merge_field.set_name(merge_field.get_name() + "_Renamed")


        doc.save(docs_base.artifacts_dir + "WorkingWithFields.rename_merge_fields.doc")
        #ExEnd:RenameMergeFields


    #ExStart:MergeField
    #
    # Represents a facade object for a merge field in a Microsoft Word document.
    #
    class MergeField :

        mFieldStart = None
        mFieldSeparator = None
        mFieldEnd = None

        gRegex = re.compile("\\s*(MERGEFIELD\\s|)(\\s|)(\\S+)\\s+")

        def __init__(self, fieldStart) :

            if fieldStart == None :
                raise ValueError("fieldStart")
            if fieldStart.field_type != aw.fields.FieldType.FIELD_MERGE_FIELD :
                raise ValueError("Field start type must be FieldMergeField.")

            self.mFieldStart = fieldStart


            # Find the field separator node.
            self.mFieldSeparator = fieldStart.get_field().separator
            if self.mFieldSeparator == None :
                raise RuntimeError("Cannot find field separator.")

            self.mFieldEnd = fieldStart.get_field().end


        #
        # Gets the name of the merge field.
        #
        def get_name(self) :
            print(type(self.mFieldStart))
            return self.mFieldStart.get_field().result.replace("«", "").replace("»", "")

        #
        # Gets the name of the merge field.
        #
        def set_name(self, name : str) :

            # Merge field name is stored in the field result which is a Run
            # node between field separator and field end.
            field_result = self.mFieldSeparator.next_sibling.as_run()
            field_result.text = f"«{name}»"

            # But sometimes the field result can consist of more than one run, delete these runs.
            self.remove_same_parent(field_result.next_sibling, self.mFieldEnd)

            self.update_field_code(name)


        def update_field_code(self, field_name : str) :

            # Field code is stored in a Run node between field start and field separator.
            field_code = self.mFieldStart.next_sibling.as_run()

            match = self.gRegex.match(self.mFieldStart.get_field().get_field_code())

            new_field_code = f" {match.group(1)}{field_name} "
            field_code.text = new_field_code

            # But sometimes the field code can consist of more than one run, delete these runs.
            self.remove_same_parent(field_code.next_sibling, self.mFieldSeparator)


        #
        # Removes nodes from start up to but not including the end node.
        # Start and end are assumed to have the same parent.
        #
        @staticmethod
        def remove_same_parent(start_node, end_node) :

            if (end_node != None and start_node.parent_node != end_node.parent_node) :
                raise ValueError("Start and end nodes are expected to have the same parent.")

            cur_child = start_node
            while (cur_child != None and cur_child != end_node) :

                next_child = cur_child.next_sibling
                cur_child.remove()
                cur_child = next_child


    #ExEnd:MergeField

    def test_remove_field(self) :

        #ExStart:RemoveField
        doc = aw.Document(docs_base.my_dir + "Various fields.docx")

        field = doc.range.fields[0]
        field.remove()
        #ExEnd:RemoveField

    def test_insert_toa_field_without_document_builder(self) :

        #ExStart:InsertTOAFieldWithoutDocumentBuilder
        doc = aw.Document()
        para = aw.Paragraph(doc)

        # We want to insert TA and TOA fields like this:
        #  TA  \c 1 \l "Value 0"
        #  TOA  \c 1

        field_ta = para.append_field(aw.fields.FieldType.FIELD_TOAENTRY, False).as_field_ta()
        field_ta.entry_category = "1"
        field_ta.long_citation = "Value 0"

        doc.first_section.body.append_child(para)

        para = aw.Paragraph(doc)

        field_toa = para.append_field(aw.fields.FieldType.FIELD_TOA, False).as_field_toa()
        field_toa.entry_category = "1"
        doc.first_section.body.append_child(para)

        field_toa.update()

        doc.save(docs_base.artifacts_dir + "WorkingWithFields.insert_toa_field_without_document_builder.docx")
        #ExEnd:InsertTOAFieldWithoutDocumentBuilder


    def test_insert_nested_fields(self) :

        #ExStart:InsertNestedFields
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        for i in range(0, 5) :
            builder.insert_break(aw.BreakType.PAGE_BREAK)

        builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)

        # We want to insert a field like this:
        #  IF PAGE <> NUMPAGES "See Next Page" "Last Page"
        field = builder.insert_field("IF ")
        builder.move_to(field.separator)
        builder.insert_field("PAGE")
        builder.write(" <> ")
        builder.insert_field("NUMPAGES")
        builder.write(" \"See Next Page\" \"Last Page\" ")

        field.update()

        doc.save(docs_base.artifacts_dir + "WorkingWithFields.insert_nested_fields.docx")
        #ExEnd:InsertNestedFields


    def test_insert_merge_field_using_dom(self) :

        #ExStart:InsertMergeFieldUsingDOM
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        para = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)[0].as_paragraph()

        builder.move_to(para)

        # We want to insert a merge field like this:
        #  " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v"

        field = builder.insert_field(aw.fields.FieldType.FIELD_MERGE_FIELD, False).as_field_merge_field()

        #  " MERGEFIELD Test1"
        field.field_name = "Test1"

        #  " MERGEFIELD Test1 \\b Test2"
        field.text_before = "Test2"

        #  " MERGEFIELD Test1 \\b Test2 \\f Test3
        field.text_after = "Test3"

        #  " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m"
        field.is_mapped = True

        #  " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v"
        field.is_vertical_formatting = True

        # Finally update this merge field
        field.update()

        doc.save(docs_base.artifacts_dir + "WorkingWithFields.insert_merge_field_using_dom.docx")
        #ExEnd:InsertMergeFieldUsingDOM


    def test_insert_mail_merge_address_block_field_using_dom(self) :

        #ExStart:InsertMailMergeAddressBlockFieldUsingDOM
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        para = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)[0].as_paragraph()

        builder.move_to(para)

        # We want to insert a mail merge address block like this:
        #  ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\"

        field = builder.insert_field(aw.fields.FieldType.FIELD_ADDRESS_BLOCK, False).as_field_address_block()

        #  ADDRESSBLOCK \\c 1"
        field.include_country_or_region_name = "1"

        #  ADDRESSBLOCK \\c 1 \\d"
        field.format_address_on_country_or_region = True

        #  ADDRESSBLOCK \\c 1 \\d \\e Test2
        field.excluded_country_or_region_name = "Test2"

        #  ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3
        field.name_and_address_format = "Test3"

        #  ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\"
        field.language_id = "Test 4"

        field.update()

        doc.save(docs_base.artifacts_dir + "WorkingWithFields.insert_mail_merge_address_block_field_using_dom.docx")
        #ExEnd:InsertMailMergeAddressBlockFieldUsingDOM


    def test_insert_field_include_text_without_document_builder(self) :

        #ExStart:InsertFieldIncludeTextWithoutDocumentBuilder
        doc = aw.Document()

        para = aw.Paragraph(doc)

        # We want to insert an INCLUDETEXT field like this:
        #  INCLUDETEXT  "file path"

        field_include_text = para.append_field(aw.fields.FieldType.FIELD_INCLUDE_TEXT, False).as_field_include_text()
        field_include_text.bookmark_name = "bookmark"
        field_include_text.source_full_name = docs_base.my_dir + "IncludeText.docx"

        doc.first_section.body.append_child(para)

        field_include_text.update()

        doc.save(docs_base.artifacts_dir + "WorkingWithFields.insert_include_field_without_document_builder.docx")
        #ExEnd:InsertFieldIncludeTextWithoutDocumentBuilder


    def test_insert_field_none(self) :

        #ExStart:InsertFieldNone
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        field = builder.insert_field(aw.fields.FieldType.FIELD_NONE, False)

        doc.save(docs_base.artifacts_dir + "WorkingWithFields.insert_field_none.docx")
        #ExEnd:InsertFieldNone


    def test_insert_field(self) :

        #ExStart:InsertField
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_field("MERGEFIELD MyFieldName \* MERGEFORMAT")

        doc.save(docs_base.artifacts_dir + "WorkingWithFields.insert_field.docx")
        #ExEnd:InsertField


    def test_insert_author_field(self) :

        #ExStart:InsertAuthorField
        doc = aw.Document()

        para = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)[0].as_paragraph()

        # We want to insert an AUTHOR field like this:
        #  AUTHOR Test1

        field = para.append_field(aw.fields.FieldType.FIELD_AUTHOR, False).as_field_author()
        field.author_name = "Test1" #  AUTHOR Test1

        field.update()

        doc.save(docs_base.artifacts_dir + "WorkingWithFields.insert_author_field.docx")
        #ExEnd:InsertAuthorField


    def test_insert_ask_field_with_out_document_builder(self) :

        #ExStart:InsertASKFieldWithOutDocumentBuilder
        doc = aw.Document()

        para = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)[0].as_paragraph()

        # We want to insert an Ask field like this:
        #  ASK \"Test 1\" Test2 \\d Test3 \\o

        field = para.append_field(aw.fields.FieldType.FIELD_ASK, False).as_field_ask()

        #  ASK \"Test 1\" "
        field.bookmark_name = "Test 1"

        #  ASK \"Test 1\" Test2
        field.prompt_text = "Test2"

        #  ASK \"Test 1\" Test2 \\d Test3
        field.default_response = "Test3"

        #  ASK \"Test 1\" Test2 \\d Test3 \\o
        field.prompt_once_on_mail_merge = True

        field.update()

        doc.save(docs_base.artifacts_dir + "WorkingWithFields.insert_ask_field_with_out_document_builder.docx")
        #ExEnd:InsertASKFieldWithOutDocumentBuilder


    def test_insert_advance_field_with_out_document_builder(self) :

        #ExStart:InsertAdvanceFieldWithOutDocumentBuilder
        doc = aw.Document()

        para = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)[0].as_paragraph()

        # We want to insert an Advance field like this:
        #  ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100

        field = para.append_field(aw.fields.FieldType.FIELD_ADVANCE, False).as_field_advance()

        #  ADVANCE \\d 10 "
        field.down_offset = "10"

        #  ADVANCE \\d 10 \\l 10
        field.left_offset = "10"

        #  ADVANCE \\d 10 \\l 10 \\r -3.3
        field.right_offset = "-3.3"

        #  ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0
        field.up_offset = "0"

        #  ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100
        field.horizontal_position = "100"

        #  ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100
        field.vertical_position = "100"

        field.update()

        doc.save(docs_base.artifacts_dir + "WorkingWithFields.insert_advance_field_with_out_document_builder.docx")
        #ExEnd:InsertAdvanceFieldWithOutDocumentBuilder


    def test_get_mail_merge_field_names(self) :

        #ExStart:GetFieldNames
        doc = aw.Document()

        field_names = doc.mail_merge.get_field_names()
        #ExEnd:GetFieldNames
        print(f"\nDocument have {len(field_names)} fields.")


    def test_mapped_data_fields(self) :

        #ExStart:MappedDataFields
        doc = aw.Document()

        doc.mail_merge.mapped_data_fields.add("MyFieldName_InDocument", "MyFieldName_InDataSource")
        #ExEnd:MappedDataFields


    def test_delete_fields(self) :

        #ExStart:DeleteFields
        doc = aw.Document()

        doc.mail_merge.delete_fields()
        #ExEnd:DeleteFields


    def test_field_display_results(self) :

        #ExStart:FieldDisplayResults
        #ExStart:UpdateDocFields
        document = aw.Document(docs_base.my_dir + "Various fields.docx")

        document.update_fields()
        #ExEnd:UpdateDocFields

        for field in document.range.fields :
            print(field.display_result)
        #ExEnd:FieldDisplayResults


    def test_evaluate_if_condition(self) :

        #ExStart:EvaluateIFCondition
        builder = aw.DocumentBuilder()

        field = builder.insert_field("IF 1 = 1", None).as_field_if()
        actual_result = field.evaluate_condition()

        print(actual_result)
        #ExEnd:EvaluateIFCondition


    def test_convert_fields_in_paragraph(self) :

        #ExStart:ConvertFieldsInParagraph
        doc = aw.Document(docs_base.my_dir + "Linked fields.docx")

        # Pass the appropriate parameters to convert all IF fields to text that are encountered only in the last
        # paragraph of the document.
        for field in doc.first_section.body.last_paragraph.range.fields :
            if field.type == aw.fields.FieldType.FIELD_IF :
                field.unlink()

        doc.save(docs_base.artifacts_dir + "WorkingWithFields.test_file.docx")
        #ExEnd:ConvertFieldsInParagraph


    def test_convert_fields_in_document(self) :

        #ExStart:ConvertFieldsInDocument
        doc = aw.Document(docs_base.my_dir + "Linked fields.docx")

        # Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to text.
        for field in doc.range.fields :
            if field.type == aw.fields.FieldType.FIELD_IF :
                field.unlink()

        # Save the document with fields transformed to disk
        doc.save(docs_base.artifacts_dir + "WorkingWithFields.convert_fields_in_document.docx")
        #ExEnd:ConvertFieldsInDocument


    def test_convert_fields_in_body(self) :

        #ExStart:ConvertFieldsInBody
        doc = aw.Document(docs_base.my_dir + "Linked fields.docx")

        # Pass the appropriate parameters to convert PAGE fields encountered to text only in the body of the first section.
        for field in doc.first_section.body.range.fields :
            if field.type == aw.fields.FieldType.FIELD_PAGE :
                field.unlink()

        doc.save(docs_base.artifacts_dir + "WorkingWithFields.convert_fields_in_body.docx")
        #ExEnd:ConvertFieldsInBody


    def test_change_locale(self) :

        #ExStart:ChangeLocale
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_field("MERGEFIELD Date")

        # Store the current culture so it can be set back once mail merge is complete.
        loc = locale.getlocale()
        # Set to German language so dates and numbers are formatted using this culture during mail merge.
        locale.setlocale(locale.LC_ALL, 'de_DE')

        doc.mail_merge.execute(["Date"], [datetime.today()] )

        locale.setlocale(locale.LC_ALL, loc)

        doc.save(docs_base.artifacts_dir + "WorkingWithFields.change_locale.docx")
        #ExEnd:ChangeLocale


if __name__ == '__main__':
    unittest.main()
