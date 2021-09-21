import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithImages(docs_base.DocsExamplesBase):

    
    def test_add_image_to_each_page(self) :
        
        doc = aw.Document(docs_base.my_dir + "Document.docx")

        # Create and attach collector before the document before page layout is built.
        layoutCollector = aw.layout.LayoutCollector(doc)

        # Images in a document are added to paragraphs to add an image to every page we need
        # to find at any paragraph belonging to each page.
        for page in range(1, doc.page_count) :
            for para in doc.first_section.body.paragraphs :
                para = para.as_paragraph()
                
                # Check if the current paragraph belongs to the target page.
                if (layoutCollector.get_start_page_index(para) == page) :
                    self.add_image_to_page(paragraph, page, docs_base.images_dir)
                    break
            

        # If we need to save the document as a PDF or image, call UpdatePageLayout() method.
        doc.update_page_layout()

        doc.save(docs_base.artifacts_dir + "WorkingWithImages.add_image_to_each_page.docx")
        

    # <summary>
    # Adds an image to a page using the supplied paragraph.
    # </summary>
    # <param name="para">The paragraph to an an image to.</param>
    # <param name="page">The page number the paragraph appears on.</param>
    @staticmethod
    def add_image_to_page(para : aw.Paragraph, page : int, imagesDir : str) :
        
        doc = para.document.as_document()

        builder = aw.DocumentBuilder(doc)
        builder.move_to(para)

        # Insert a logo to the top left of the page to place it in front of all other text.
        builder.insert_image(docs_base.images_dir + "Transparent background logo.png", aw.drawing.RelativeHorizontalPosition.PAGE, 60,
            aw.drawing.RelativeVerticalPosition.PAGE, 60, -1, -1, aw.drawing.WrapType.NONE)

        # Insert a textbox next to the image which contains some text consisting of the page number.
        textBox = aw.drawing.Shape(doc, aw.drawing.ShapeType.TEXT_BOX)

        # We want a floating shape relative to the page.
        textBox.wrap_type = aw.drawing.WrapType.NONE
        textBox.relative_horizontal_position = aw.drawing.RelativeHorizontalPosition.PAGE
        textBox.relative_vertical_position = aw.drawing.RelativeVerticalPosition.PAGE

        textBox.height = 30
        textBox.width = 200
        textBox.left = 150
        textBox.top = 80

        textBox.append_child(aw.Paragraph(doc))
        builder.insert_node(textBox)
        builder.move_to(textBox.first_child)
        builder.writeln("This is a custom note for page " + page)
        

    def test_insert_barcode_image(self) :
        
        #ExStart:InsertBarcodeImage
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # The number of pages the document should have
        numPages = 4
        # The document starts with one section, insert the barcode into this existing section
        self.insert_barcode_into_footer(builder, doc.first_section, aw.HeaderFooterType.FOOTER_PRIMARY)

        for i in range(1, numPages) :
            
            # Clone the first section and add it into the end of the document
            cloneSection = doc.first_section.clone(False).as_section()
            cloneSection.page_setup.section_start = aw.SectionStart.NEW_PAGE
            doc.append_child(cloneSection)

            # Insert the barcode and other information into the footer of the section
            self.insert_barcode_into_footer(builder, cloneSection, aw.HeaderFooterType.FOOTER_PRIMARY)
            

        # Save the document as a PDF to disk
        # You can also save this directly to a stream
        doc.save(docs_base.artifacts_dir + "InsertBarcodeImage.docx")
        #ExEnd:InsertBarcodeImage
        

    #ExStart:InsertBarcodeIntoFooter
    @staticmethod
    def insert_barcode_into_footer(builder : aw.DocumentBuilder, section : aw.Section, footerType : aw.HeaderFooterType) :
        
        # Move to the footer type in the specific section.
        builder.move_to_section(section.document.index_of(section))
        builder.move_to_header_footer(footerType)

        # Insert the barcode, then move to the next line and insert the ID along with the page number.
        # Use pageId if you need to insert a different barcode on each page. 0 = First page, 1 = Second page etc.
        builder.insert_image(docs_base.images_dir + "Barcode.png")
        builder.writeln()
        builder.write("1234567890")
        builder.insert_field("PAGE")

        # Create a right-aligned tab at the right margin.
        tabPos = section.page_setup.page_width - section.page_setup.right_margin - section.page_setup.left_margin
        builder.current_paragraph.paragraph_format.tab_stops.add(aw.TabStop(tabPos, aw.TabAlignment.RIGHT, aw.TabLeader.NONE))

        # Move to the right-hand side of the page and insert the page and page total.
        builder.write(aw.ControlChar.TAB)
        builder.insert_field("PAGE")
        builder.write(" of ")
        builder.insert_field("NUMPAGES")
        
    #ExEnd:InsertBarcodeIntoFooter


if __name__ == '__main__':
    unittest.main()