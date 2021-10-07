import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base
import extract_content_helper as helper

import aspose.words as aw

class ExtractContent(docs_base.DocsExamplesBase):

    def test_extract_content_between_block_level_nodes(self) :
        
        #ExStart:ExtractContentBetweenBlockLevelNodes
        doc = aw.Document(docs_base.my_dir + "Extract content.docx")

        startPara = doc.last_section.get_child(aw.NodeType.PARAGRAPH, 2, True).as_paragraph()
        endTable = doc.last_section.get_child(aw.NodeType.TABLE, 0, True).as_table()

        # Extract the content between these nodes in the document. Include these markers in the extraction.
        extractedNodes = helper.ExtractContentHelper.extract_content(startPara, endTable, True)

        # Let's reverse the array to make inserting the content back into the document easier.
        extractedNodes.reverse()

        while (len(extractedNodes) > 0) :
            
            # Insert the last node from the reversed list.
            endTable.parent_node.insert_after(extractedNodes[0], endTable)
            # Remove this node from the list after insertion.
            del extractedNodes[0]
            

        doc.save(docs_base.artifacts_dir + "ExtractContent.extract_content_between_block_level_nodes.docx")
        #ExEnd:ExtractContentBetweenBlockLevelNodes
        

    def test_extract_content_between_bookmark(self) :
        
        #ExStart:ExtractContentBetweenBookmark
        doc = aw.Document(docs_base.my_dir + "Extract content.docx")

        section = doc.sections[0]
        section.page_setup.left_margin = 70.85

        # Retrieve the bookmark from the document.
        bookmark = doc.range.bookmarks.get_by_name("Bookmark1")
        # We use the BookmarkStart and BookmarkEnd nodes as markers.
        bookmarkStart = bookmark.bookmark_start
        bookmarkEnd = bookmark.bookmark_end

        # Firstly, extract the content between these nodes, including the bookmark.
        extractedNodesInclusive = helper.ExtractContentHelper.extract_content(bookmarkStart, bookmarkEnd, True)
            
        dstDoc = helper.ExtractContentHelper.generate_document(doc, extractedNodesInclusive)
        dstDoc.save(docs_base.artifacts_dir + "ExtractContent.extract_content_between_bookmark.including_bookmark.docx")

        # Secondly, extract the content between these nodes this time without including the bookmark.
        extractedNodesExclusive = helper.ExtractContentHelper.extract_content(bookmarkStart, bookmarkEnd, False)
            
        dstDoc = helper.ExtractContentHelper.generate_document(doc, extractedNodesExclusive)
        dstDoc.save(docs_base.artifacts_dir + "ExtractContent.extract_content_between_bookmark.without_bookmark.docx")
        #ExEnd:ExtractContentBetweenBookmark
        

    def test_extract_content_between_comment_range(self) :
        
        #ExStart:ExtractContentBetweenCommentRange
        doc = aw.Document(docs_base.my_dir + "Extract content.docx")

        # This is a quick way of getting both comment nodes.
        # Your code should have a proper method of retrieving each corresponding start and end node.
        commentStart = doc.get_child(aw.NodeType.COMMENT_RANGE_START, 0, True).as_comment_range_start()
        commentEnd = doc.get_child(aw.NodeType.COMMENT_RANGE_END, 0, True).as_comment_range_end()

        # Firstly, extract the content between these nodes including the comment as well.
        extractedNodesInclusive = helper.ExtractContentHelper.extract_content(commentStart, commentEnd, True)
            
        dstDoc = helper.ExtractContentHelper.generate_document(doc, extractedNodesInclusive)
        dstDoc.save(docs_base.artifacts_dir + "ExtractContent.extract_content_between_comment_range.including_comment.docx")

        # Secondly, extract the content between these nodes without the comment.
        extractedNodesExclusive = helper.ExtractContentHelper.extract_content(commentStart, commentEnd, False)
            
        dstDoc = helper.ExtractContentHelper.generate_document(doc, extractedNodesExclusive)
        dstDoc.save(docs_base.artifacts_dir + "ExtractContent.extract_content_between_comment_range.without_comment.docx")
        #ExEnd:ExtractContentBetweenCommentRange
        

    def test_extract_content_between_paragraphs(self) :
        
        #ExStart:ExtractContentBetweenParagraphs
        doc = aw.Document(docs_base.my_dir + "Extract content.docx")

        startPara = doc.first_section.body.get_child(aw.NodeType.PARAGRAPH, 6, True).as_paragraph()
        endPara = doc.first_section.body.get_child(aw.NodeType.PARAGRAPH, 10, True).as_paragraph()

        # Extract the content between these nodes in the document. Include these markers in the extraction.
        extractedNodes = helper.ExtractContentHelper.extract_content(startPara, endPara, True)

        dstDoc = helper.ExtractContentHelper.generate_document(doc, extractedNodes)
        dstDoc.save(docs_base.artifacts_dir + "ExtractContent.extract_content_between_paragraphs.docx")
        #ExEnd:ExtractContentBetweenParagraphs
        

    def test_extract_content_between_paragraph_styles(self) :
        
        #ExStart:ExtractContentBetweenParagraphStyles
        doc = aw.Document(docs_base.my_dir + "Extract content.docx")

        # Gather a list of the paragraphs using the respective heading styles.
        parasStyleHeading1 = helper.ExtractContentHelper.paragraphs_by_style_name(doc, "Heading 1")
        parasStyleHeading3 = helper.ExtractContentHelper.paragraphs_by_style_name(doc, "Heading 3")

        # Use the first instance of the paragraphs with those styles.
        startPara1 = parasStyleHeading1[0]
        endPara1 = parasStyleHeading3[0]

        # Extract the content between these nodes in the document. Don't include these markers in the extraction.
        extractedNodes = helper.ExtractContentHelper.extract_content(startPara1, endPara1, False)

        dstDoc = helper.ExtractContentHelper.generate_document(doc, extractedNodes)
        dstDoc.save(docs_base.artifacts_dir + "ExtractContent.extract_content_between_paragraph_styles.docx")
        #ExEnd:ExtractContentBetweenParagraphStyles
        

    def test_extract_content_between_runs(self) :
        
        #ExStart:ExtractContentBetweenRuns
        doc = aw.Document(docs_base.my_dir + "Extract content.docx")

        para = doc.get_child(aw.NodeType.PARAGRAPH, 7, True).as_paragraph()

        startRun = para.runs[1]
        endRun = para.runs[4]

        # Extract the content between these nodes in the document. Include these markers in the extraction.
        extractedNodes = helper.ExtractContentHelper.extract_content(startRun, endRun, True)

        node = extractedNodes[0]
        print(node.to_string(aw.SaveFormat.TEXT))
        #ExEnd:ExtractContentBetweenRuns
        

    def test_extract_content_using_field(self) :
        
        #ExStart:ExtractContentUsingField
        doc = aw.Document(docs_base.my_dir + "Extract content.docx")
        builder = aw.DocumentBuilder(doc)

        # Pass the first boolean parameter to get the DocumentBuilder to move to the FieldStart of the field.
        # We could also get FieldStarts of a field using GetChildNode method as in the other examples.
        builder.move_to_merge_field("Fullname", False, False)

        # The builder cursor should be positioned at the start of the field.
        startField = builder.current_node.as_field_start()
        endPara = doc.first_section.get_child(aw.NodeType.PARAGRAPH, 5, True).as_paragraph()

        # Extract the content between these nodes in the document. Don't include these markers in the extraction.
        extractedNodes = helper.ExtractContentHelper.extract_content(startField, endPara, False)

        dstDoc = helper.ExtractContentHelper.generate_document(doc, extractedNodes)
        dstDoc.save(docs_base.artifacts_dir + "ExtractContent.extract_content_using_field.docx")
        #ExEnd:ExtractContentUsingField
        

    def test_extract_table_of_contents(self) :
        
        #ExStart:ExtractTableOfContents
        doc = aw.Document(docs_base.my_dir + "Table of contents.docx")

        for field in doc.range.fields :
            
            if (field.type == aw.fields.FieldType.FIELD_HYPERLINK) :
                
                hyperlink = field.as_field_hyperlink()
                if (hyperlink.sub_address != None and hyperlink.sub_address.find("_Toc") == 0) :
                    
                    tocItem = field.start.get_ancestor(aw.NodeType.PARAGRAPH).as_paragraph()
                        
                    print(tocItem.to_string(aw.SaveFormat.TEXT).strip())
                    print("------------------")

                    bm = doc.range.bookmarks.get_by_name(hyperlink.sub_address)
                    pointer = bm.bookmark_start.get_ancestor(aw.NodeType.PARAGRAPH).as_paragraph()
                        
                    print(pointer.to_string(aw.SaveFormat.TEXT))
        #ExEnd:ExtractTableOfContents
                    
                
    def test_extract_text_only(self) :
        
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
        

    def test_extract_content_based_on_styles(self) :
        
        #ExStart:ExtractContentBasedOnStyles
        doc = aw.Document(docs_base.my_dir + "Styles.docx")

        paraStyle = "Heading 1"
        runStyle = "Intense Emphasis"

        paragraphs = ExtractContent.paragraphs_by_style_name(doc, paraStyle)
        print(f"Paragraphs with \"{paraStyle}\" styles ({len(paragraphs)}):")
            
        for paragraph in paragraphs :
            print(paragraph.to_string(aw.SaveFormat.TEXT))

        runs = ExtractContent.runs_by_style_name(doc, runStyle)
        print(f"\nRuns with \"{runStyle}\" styles ({len(runs)}):")
            
        for run in runs :
            print(run.range.text)
        #ExEnd:ExtractContentBasedOnStyles
        

    #ExStart:ParagraphsByStyleName
    @staticmethod
    def paragraphs_by_style_name(doc : aw.Document, styleName : str) :
        
        paragraphsWithStyle = []
        paragraphs = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
            
        for paragraph in paragraphs :
            paragraph = paragraph.as_paragraph()
            if (paragraph.paragraph_format.style.name == styleName) :
                paragraphsWithStyle.append(paragraph)

        return paragraphsWithStyle
        
    #ExEnd:ParagraphsByStyleName
        
    #ExStart:RunsByStyleName
    @staticmethod
    def runs_by_style_name(doc : aw.Document, styleName : str) :
        
        runsWithStyle = []
        runs = doc.get_child_nodes(aw.NodeType.RUN, True)
            
        for run in runs :
            run = run.as_run()
            if (run.font.style.name == styleName) :
                runsWithStyle.append(run)

        return runsWithStyle
        
    #ExEnd:RunsByStyleName

    def test_extract_print_text(self) :
        
        #ExStart:ExtractText
        doc = aw.Document(docs_base.my_dir + "Tables.docx")

            
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
        

    def test_extract_images_to_files(self) :
        
        #ExStart:ExtractImagesToFiles
        doc = aw.Document(docs_base.my_dir + "Images.docx")

        shapes = doc.get_child_nodes(aw.NodeType.SHAPE, True)
        imageIndex = 0
            
        for shape in shapes :
            shape = shape.as_shape()
            if (shape.has_image) :
                
                imageFileName = f"Image.ExportImages.{imageIndex}_{aw.FileFormatUtil.image_type_to_extension(shape.image_data.image_type)}"

                shape.image_data.save(docs_base.artifacts_dir + imageFileName)
                imageIndex += 1
                
        #ExEnd:ExtractImagesToFiles
        
    

if __name__ == '__main__':
    unittest.main()