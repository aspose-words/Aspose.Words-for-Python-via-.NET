import unittest
import os
import sys
import uuid

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw
import aspose.pydrawing as drawing

class WorkingWithSdt(docs_base.DocsExamplesBase):

        def test_check_box_type_content_control(self) :

            #ExStart:CheckBoxTypeContentControl
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc)

            sdt_check_box = aw.markup.StructuredDocumentTag(doc, aw.markup.SdtType.CHECKBOX, aw.markup.MarkupLevel.INLINE)
            builder.insert_node(sdt_check_box)

            doc.save(docs_base.artifacts_dir + "WorkingWithSdt.check_box_type_content_control.docx", aw.SaveFormat.DOCX)
            #ExEnd:CheckBoxTypeContentControl


        def test_current_state_of_check_box(self) :

            #ExStart:SetCurrentStateOfCheckBox
            doc = aw.Document(docs_base.my_dir + "Structured document tags.docx")

            # Get the first content control from the document.
            sdt_check_box = doc.get_child(aw.NodeType.STRUCTURED_DOCUMENT_TAG, 0, True).as_structured_document_tag()

            if (sdt_check_box.sdt_type == aw.markup.SdtType.CHECKBOX) :
                sdt_check_box.checked = True

            doc.save(docs_base.artifacts_dir + "WorkingWithSdt.current_state_of_check_box.docx")
            #ExEnd:SetCurrentStateOfCheckBox


        def test_modify_content_controls(self) :

            #ExStart:ModifyContentControls
            doc = aw.Document(docs_base.my_dir + "Structured document tags.docx")

            for sdt in doc.get_child_nodes(aw.NodeType.STRUCTURED_DOCUMENT_TAG, True) :
                sdt = sdt.as_structured_document_tag()

                if (sdt.sdt_type == aw.markup.SdtType.PLAIN_TEXT) :

                    sdt.remove_all_children()
                    para = sdt.append_child(aw.Paragraph(doc)).as_paragraph()
                    run = aw.Run(doc, "new text goes here")
                    para.append_child(run)

                elif (sdt.sdt_type == aw.markup.SdtType.DROP_DOWN_LIST) :

                    second_item = sdt.list_items[2]
                    sdt.list_items.selected_value = second_item

                elif (sdt.sdt_type == aw.markup.SdtType.PICTURE) :

                    shape = sdt.get_child(NodeType.shape, 0, True).as_shape()
                    if (shape.has_image) :
                        shape.image_data.set_image(docs_base.images_dir + "Watermark.png")

            doc.save(docs_base.artifacts_dir + "WorkingWithSdt.modify_content_controls.docx")
            #ExEnd:ModifyContentControls


        def test_combo_box_content_control(self) :

            #ExStart:ComboBoxContentControl
            doc = aw.Document()

            sdt = aw.markup.StructuredDocumentTag(doc, aw.markup.SdtType.COMBO_BOX, aw.markup.MarkupLevel.BLOCK)
            sdt.list_items.add(aw.markup.SdtListItem("Choose an item", "-1"))
            sdt.list_items.add(aw.markup.SdtListItem("Item 1", "1"))
            sdt.list_items.add(aw.markup.SdtListItem("Item 2", "2"))
            doc.first_section.body.append_child(sdt)

            doc.save(docs_base.artifacts_dir + "WorkingWithSdt.combo_box_content_control.docx")
            #ExEnd:ComboBoxContentControl


        def test_rich_text_box_content_control(self) :

            #ExStart:RichTextBoxContentControl
            doc = aw.Document()

            sdt_rich_text = aw.markup.StructuredDocumentTag(doc, aw.markup.SdtType.RICH_TEXT, aw.markup.MarkupLevel.BLOCK)

            para = aw.Paragraph(doc)
            run = aw.Run(doc)
            run.text = "Hello World"
            run.font.color = drawing.Color.green
            para.runs.add(run)
            sdt_rich_text.child_nodes.add(para)
            doc.first_section.body.append_child(sdt_rich_text)

            doc.save(docs_base.artifacts_dir + "WorkingWithSdt.rich_text_box_content_control.docx")
            #ExEnd:RichTextBoxContentControl


        def test_set_content_control_color(self) :

            #ExStart:SetContentControlColor
            doc = aw.Document(docs_base.my_dir + "Structured document tags.docx")

            sdt = doc.get_child(aw.NodeType.STRUCTURED_DOCUMENT_TAG, 0, True).as_structured_document_tag()
            sdt.color = drawing.Color.red

            doc.save(docs_base.artifacts_dir + "WorkingWithSdt.set_content_control_color.docx")
            #ExEnd:SetContentControlColor


        def test_clear_contents_control(self) :

            #ExStart:ClearContentsControl
            doc = aw.Document(docs_base.my_dir + "Structured document tags.docx")

            sdt = doc.get_child(aw.NodeType.STRUCTURED_DOCUMENT_TAG, 0, True).as_structured_document_tag()
            sdt.clear()

            doc.save(docs_base.artifacts_dir + "WorkingWithSdt.clear_contents_control.doc")
            #ExEnd:ClearContentsControl


        def test_bind_sd_tto_custom_xml_part(self) :

            #ExStart:BindSDTtoCustomXmlPart
            doc = aw.Document()
            xml_part = doc.custom_xml_parts.add(str(uuid.uuid4()), "<root><text>Hello, World!</text></root>")

            sdt = aw.markup.StructuredDocumentTag(doc, aw.markup.SdtType.PLAIN_TEXT, aw.markup.MarkupLevel.BLOCK)
            doc.first_section.body.append_child(sdt)

            sdt.xml_mapping.set_mapping(xml_part, "/root[1]/text[1]", "")

            doc.save(docs_base.artifacts_dir + "WorkingWithSdt.bind_sd_tto_custom_xml_part.doc")
            #ExEnd:BindSDTtoCustomXmlPart


        def test_set_content_control_style(self) :

            #ExStart:SetContentControlStyle
            doc = aw.Document(docs_base.my_dir + "Structured document tags.docx")

            sdt = doc.get_child(aw.NodeType.STRUCTURED_DOCUMENT_TAG, 0, True).as_structured_document_tag()
            style = doc.styles.get_by_style_identifier(aw.StyleIdentifier.QUOTE)
            sdt.style = style

            doc.save(docs_base.artifacts_dir + "WorkingWithSdt.set_content_control_style.docx")
            #ExEnd:SetContentControlStyle


        def test_creating_table_repeating_section_mapped_to_custom_xml_part(self) :

            #ExStart:CreatingTableRepeatingSectionMappedToCustomXmlPart
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc)

            xml_part = doc.custom_xml_parts.add("Books",
                "<books><book><title>Everyday Italian</title><author>Giada De Laurentiis</author></book>" +
                "<book><title>Harry Potter</title><author>J K. Rowling</author></book>" +
                "<book><title>Learning XML</title><author>Erik T. Ray</author></book></books>")

            table = builder.start_table()

            builder.insert_cell()
            builder.write("Title")

            builder.insert_cell()
            builder.write("Author")

            builder.end_row()
            builder.end_table()

            repeating_section_sdt =aw.markup.StructuredDocumentTag(doc, aw.markup.SdtType.REPEATING_SECTION, aw.markup.MarkupLevel.ROW)
            repeating_section_sdt.xml_mapping.set_mapping(xml_part, "/books[1]/book", "")
            table.append_child(repeating_section_sdt)

            repeating_section_item_sdt = aw.markup.StructuredDocumentTag(doc, aw.markup.SdtType.REPEATING_SECTION_ITEM, aw.markup.MarkupLevel.ROW)
            repeating_section_sdt.append_child(repeating_section_item_sdt)

            row = aw.tables.Row(doc)
            repeating_section_item_sdt.append_child(row)

            title_sdt = aw.markup.StructuredDocumentTag(doc, aw.markup.SdtType.PLAIN_TEXT, aw.markup.MarkupLevel.CELL)
            title_sdt.xml_mapping.set_mapping(xml_part, "/books[1]/book[1]/title[1]", "")
            row.append_child(title_sdt)

            author_sdt = aw.markup.StructuredDocumentTag(doc, aw.markup.SdtType.PLAIN_TEXT, aw.markup.MarkupLevel.CELL)
            author_sdt.xml_mapping.set_mapping(xml_part, "/books[1]/book[1]/author[1]", "")
            row.append_child(author_sdt)

            doc.save(docs_base.artifacts_dir + "WorkingWithSdt.creating_table_repeating_section_mapped_to_custom_xml_part.docx")
            #ExEnd:CreatingTableRepeatingSectionMappedToCustomXmlPart


        def test_multi_section(self) :

            #ExStart:MultiSectionSDT
            doc = aw.Document(docs_base.my_dir + "Multi-section structured document tags.docx")

            tags = doc.get_child_nodes(aw.NodeType.STRUCTURED_DOCUMENT_TAG_RANGE_START, True)

            for tag in tags :
                print(tag.as_structured_document_tag_range_start().title)
            #ExEnd:MultiSectionSDT


        def test_structured_document_tag_range_start_xml_mapping(self) :

            #ExStart:StructuredDocumentTagRangeStartXmlMapping
            doc = aw.Document(docs_base.my_dir + "Multi-section structured document tags.docx")

            # Construct an XML part that contains data and add it to the document's CustomXmlPart collection.
            xml_part_id = str(uuid.uuid4())
            xml_part_content = "<root><text>Text element #1</text><text>Text element #2</text></root>"
            xml_part = doc.custom_xml_parts.add(xml_part_id, xml_part_content)
            print(xml_part.data.decode("utf-8"))

            # Create a StructuredDocumentTag that will display the contents of our CustomXmlPart in the document.
            sdt_range_start = doc.get_child(aw.NodeType.STRUCTURED_DOCUMENT_TAG_RANGE_START, 0, True).as_structured_document_tag_range_start()

            # If we set a mapping for our StructuredDocumentTag,
            # it will only display a part of the CustomXmlPart that the XPath points to.
            # This XPath will point to the contents second "<text>" element of the first "<root>" element of our CustomXmlPart.
            sdt_range_start.xml_mapping.set_mapping(xml_part, "/root[1]/text[2]", None)

            doc.save(docs_base.artifacts_dir + "WorkingWithSdt.structured_document_tag_range_start_xml_mapping.docx")
            #ExEnd:StructuredDocumentTagRangeStartXmlMapping


if __name__ == '__main__':
        unittest.main()
