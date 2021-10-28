import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR
from . import extract_content_helper as helper

class ExtractContent(DocsExamplesBase):

    def test_extract_content_between_block_level_nodes(self):

        #ExStart:ExtractContentBetweenBlockLevelNodes
        doc = aw.Document(MY_DIR + "Extract content.docx")

        start_para = doc.last_section.get_child(aw.NodeType.PARAGRAPH, 2, True).as_paragraph()
        end_table = doc.last_section.get_child(aw.NodeType.TABLE, 0, True).as_table()

        # Extract the content between these nodes in the document. Include these markers in the extraction.
        extracted_nodes = helper.ExtractContentHelper.extract_content(start_para, end_table, True)

        # Let's reverse the array to make inserting the content back into the document easier.
        extracted_nodes.reverse()

        while len(extracted_nodes) > 0:
            # Insert the last node from the reversed list.
            end_table.parent_node.insert_after(extracted_nodes[0], end_table)
            # Remove this node from the list after insertion.
            del extracted_nodes[0]

        doc.save(ARTIFACTS_DIR + "ExtractContent.extract_content_between_block_level_nodes.docx")
        #ExEnd:ExtractContentBetweenBlockLevelNodes

    def test_extract_content_between_bookmark(self):

        #ExStart:ExtractContentBetweenBookmark
        doc = aw.Document(MY_DIR + "Extract content.docx")

        section = doc.sections[0]
        section.page_setup.left_margin = 70.85

        # Retrieve the bookmark from the document.
        bookmark = doc.range.bookmarks.get_by_name("Bookmark1")
        # We use the BookmarkStart and BookmarkEnd nodes as markers.
        bookmark_start = bookmark.bookmark_start
        bookmark_end = bookmark.bookmark_end

        # Firstly, extract the content between these nodes, including the bookmark.
        extracted_nodes_inclusive = helper.ExtractContentHelper.extract_content(bookmark_start, bookmark_end, True)

        dst_doc = helper.ExtractContentHelper.generate_document(doc, extracted_nodes_inclusive)
        dst_doc.save(ARTIFACTS_DIR + "ExtractContent.extract_content_between_bookmark.including_bookmark.docx")

        # Secondly, extract the content between these nodes this time without including the bookmark.
        extracted_nodes_exclusive = helper.ExtractContentHelper.extract_content(bookmark_start, bookmark_end, False)

        dst_doc = helper.ExtractContentHelper.generate_document(doc, extracted_nodes_exclusive)
        dst_doc.save(ARTIFACTS_DIR + "ExtractContent.extract_content_between_bookmark.without_bookmark.docx")
        #ExEnd:ExtractContentBetweenBookmark

    def test_extract_content_between_comment_range(self):

        #ExStart:ExtractContentBetweenCommentRange
        doc = aw.Document(MY_DIR + "Extract content.docx")

        # This is a quick way of getting both comment nodes.
        # Your code should have a proper method of retrieving each corresponding start and end node.
        comment_start = doc.get_child(aw.NodeType.COMMENT_RANGE_START, 0, True).as_comment_range_start()
        comment_end = doc.get_child(aw.NodeType.COMMENT_RANGE_END, 0, True).as_comment_range_end()

        # Firstly, extract the content between these nodes including the comment as well.
        extracted_nodes_inclusive = helper.ExtractContentHelper.extract_content(comment_start, comment_end, True)

        dst_doc = helper.ExtractContentHelper.generate_document(doc, extracted_nodes_inclusive)
        dst_doc.save(ARTIFACTS_DIR + "ExtractContent.extract_content_between_comment_range.including_comment.docx")

        # Secondly, extract the content between these nodes without the comment.
        extracted_nodes_exclusive = helper.ExtractContentHelper.extract_content(comment_start, comment_end, False)

        dst_doc = helper.ExtractContentHelper.generate_document(doc, extracted_nodes_exclusive)
        dst_doc.save(ARTIFACTS_DIR + "ExtractContent.extract_content_between_comment_range.without_comment.docx")
        #ExEnd:ExtractContentBetweenCommentRange

    def test_extract_content_between_paragraphs(self):

        #ExStart:ExtractContentBetweenParagraphs
        doc = aw.Document(MY_DIR + "Extract content.docx")

        start_para = doc.first_section.body.get_child(aw.NodeType.PARAGRAPH, 6, True).as_paragraph()
        end_para = doc.first_section.body.get_child(aw.NodeType.PARAGRAPH, 10, True).as_paragraph()

        # Extract the content between these nodes in the document. Include these markers in the extraction.
        extracted_nodes = helper.ExtractContentHelper.extract_content(start_para, end_para, True)

        dst_doc = helper.ExtractContentHelper.generate_document(doc, extracted_nodes)
        dst_doc.save(ARTIFACTS_DIR + "ExtractContent.extract_content_between_paragraphs.docx")
        #ExEnd:ExtractContentBetweenParagraphs

    def test_extract_content_between_paragraph_styles(self):

        #ExStart:ExtractContentBetweenParagraphStyles
        doc = aw.Document(MY_DIR + "Extract content.docx")

        # Gather a list of the paragraphs using the respective heading styles.
        paras_style_heading1 = helper.ExtractContentHelper.paragraphs_by_style_name(doc, "Heading 1")
        paras_style_heading3 = helper.ExtractContentHelper.paragraphs_by_style_name(doc, "Heading 3")

        # Use the first instance of the paragraphs with those styles.
        start_para1 = paras_style_heading1[0]
        end_para1 = paras_style_heading3[0]

        # Extract the content between these nodes in the document. Don't include these markers in the extraction.
        extracted_nodes = helper.ExtractContentHelper.extract_content(start_para1, end_para1, False)

        dst_doc = helper.ExtractContentHelper.generate_document(doc, extracted_nodes)
        dst_doc.save(ARTIFACTS_DIR + "ExtractContent.extract_content_between_paragraph_styles.docx")
        #ExEnd:ExtractContentBetweenParagraphStyles

    def test_extract_content_between_runs(self):

        #ExStart:ExtractContentBetweenRuns
        doc = aw.Document(MY_DIR + "Extract content.docx")

        para = doc.get_child(aw.NodeType.PARAGRAPH, 7, True).as_paragraph()

        start_run = para.runs[1]
        end_run = para.runs[4]

        # Extract the content between these nodes in the document. Include these markers in the extraction.
        extracted_nodes = helper.ExtractContentHelper.extract_content(start_run, end_run, True)

        node = extracted_nodes[0]
        print(node.to_string(aw.SaveFormat.TEXT))
        #ExEnd:ExtractContentBetweenRuns

    def test_extract_content_using_field(self):

        #ExStart:ExtractContentUsingField
        doc = aw.Document(MY_DIR + "Extract content.docx")
        builder = aw.DocumentBuilder(doc)

        # Pass the first boolean parameter to get the DocumentBuilder to move to the FieldStart of the field.
        # We could also get FieldStarts of a field using GetChildNode method as in the other examples.
        builder.move_to_merge_field("Fullname", False, False)

        # The builder cursor should be positioned at the start of the field.
        start_field = builder.current_node.as_field_start()
        end_para = doc.first_section.get_child(aw.NodeType.PARAGRAPH, 5, True).as_paragraph()

        # Extract the content between these nodes in the document. Don't include these markers in the extraction.
        extracted_nodes = helper.ExtractContentHelper.extract_content(start_field, end_para, False)

        dst_doc = helper.ExtractContentHelper.generate_document(doc, extracted_nodes)
        dst_doc.save(ARTIFACTS_DIR + "ExtractContent.extract_content_using_field.docx")
        #ExEnd:ExtractContentUsingField

    def test_extract_table_of_contents(self):

        #ExStart:ExtractTableOfContents
        doc = aw.Document(MY_DIR + "Table of contents.docx")

        for field in doc.range.fields:
            if field.type == aw.fields.FieldType.FIELD_HYPERLINK:
                hyperlink = field.as_field_hyperlink()
                if hyperlink.sub_address is not None and "_Toc" not in hyperlink.sub_address:
                    toc_item = field.start.get_ancestor(aw.NodeType.PARAGRAPH).as_paragraph()

                    print(toc_item.to_string(aw.SaveFormat.TEXT).strip())
                    print("------------------")

                    bookmark = doc.range.bookmarks.get_by_name(hyperlink.sub_address)
                    pointer = bookmark.bookmark_start.get_ancestor(aw.NodeType.PARAGRAPH).as_paragraph()

                    print(pointer.to_string(aw.SaveFormat.TEXT))
        #ExEnd:ExtractTableOfContents

    def test_extract_text_only(self):

        #ExStart:ExtractTextOnly
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_field("MERGEFIELD Field")

        print("GetText() Result: " + doc.get_text())

        # When converted to text it will not retrieve fields code or special characters,
        # but will still contain some natural formatting characters such as paragraph markers etc.
        # This is the same as "viewing" the document as if it was opened in a text editor.
        print("ToString() Result: " + doc.to_string(aw.SaveFormat.TEXT))
        #ExEnd:ExtractTextOnly

    def test_extract_content_based_on_styles(self):

        #ExStart:ExtractContentBasedOnStyles
        doc = aw.Document(MY_DIR + "Styles.docx")

        para_style = "Heading 1"
        run_style = "Intense Emphasis"

        paragraphs = ExtractContent.paragraphs_by_style_name(doc, para_style)
        print(f'Paragraphs with "{para_style}" styles ({len(paragraphs)}):')

        for paragraph in paragraphs:
            print(paragraph.to_string(aw.SaveFormat.TEXT))

        runs = ExtractContent.runs_by_style_name(doc, run_style)
        print(f'\nRuns with "{run_style}" styles ({len(runs)}):')

        for run in runs:
            print(run.range.text)
        #ExEnd:ExtractContentBasedOnStyles


    #ExStart:ParagraphsByStyleName
    @staticmethod
    def paragraphs_by_style_name(doc: aw.Document, style_name: str):

        paragraphs_with_style = []
        paragraphs = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)

        for paragraph in paragraphs:
            paragraph = paragraph.as_paragraph()
            if paragraph.paragraph_format.style.name == style_name:
                paragraphs_with_style.append(paragraph)

        return paragraphs_with_style

    #ExEnd:ParagraphsByStyleName

    #ExStart:RunsByStyleName
    @staticmethod
    def runs_by_style_name(doc: aw.Document, style_name: str):

        runs_with_style = []
        runs = doc.get_child_nodes(aw.NodeType.RUN, True)

        for run in runs:
            run = run.as_run()
            if run.font.style.name == style_name:
                runs_with_style.append(run)

        return runs_with_style

    #ExEnd:RunsByStyleName

    def test_extract_print_text(self):

        #ExStart:ExtractText
        doc = aw.Document(MY_DIR + "Tables.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()

        # The range text will include control characters such as "\a" for a cell.
        # You can call ToString and pass SaveFormat.text on the desired node to find the plain text content.

        print("Contents of the table: ")
        print(table.range.text)
        #ExEnd:ExtractText

        #ExStart:PrintTextRangeOFRowAndTable
        print("\nContents of the row: ")
        print(table.rows[1].range.text)

        print("\nContents of the cell: ")
        print(table.last_row.last_cell.range.text)
        #ExEnd:PrintTextRangeOFRowAndTable

    def test_extract_images_to_files(self):

        #ExStart:ExtractImagesToFiles
        doc = aw.Document(MY_DIR + "Images.docx")

        shapes = doc.get_child_nodes(aw.NodeType.SHAPE, True)
        image_index = 0

        for shape in shapes:
            shape = shape.as_shape()
            if shape.has_image:
                image_file_name = f"Image.ExportImages.{image_index}_{aw.FileFormatUtil.image_type_to_extension(shape.image_data.image_type)}"

                shape.image_data.save(ARTIFACTS_DIR + image_file_name)
                image_index += 1

        #ExEnd:ExtractImagesToFiles
