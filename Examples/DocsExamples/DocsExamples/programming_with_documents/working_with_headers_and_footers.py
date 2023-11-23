import aspose.words as aw
from docs_examples_base import DocsExamplesBase, ARTIFACTS_DIR, IMAGES_DIR

class WorkingWithHeadersAndFooters(DocsExamplesBase):

    def test_create_header_footer(self):
        #ExStart:CreateHeaderFooter
        #GistId:2e1b2b28253780881d116e3a873ee668
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Use HeaderPrimary and FooterPrimary
        # if you want to set header/footer for all document.
        # This header/footer type also responsible for odd pages.
        #ExStart:HeaderFooterType
        #GistId:2e1b2b28253780881d116e3a873ee668
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.write("Header for page.")
        #ExEnd:HeaderFooterType

        builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)
        builder.write("Footer for page.")

        doc.save(ARTIFACTS_DIR + "WorkingWithHeadersAndFooters.create_header_footer.docx")
        #ExEnd:CreateHeaderFooter

    def test_different_first_page(self):
        #ExStart:DifferentFirstPage
        #GistId:2e1b2b28253780881d116e3a873ee668
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.page_setup.different_first_page_header_footer = True

        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_FIRST)
        builder.write("Header for the first page.")
        builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_FIRST)
        builder.write("Footer for the first page.")

        builder.move_to_section(0)
        builder.writeln("Page 1")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln("Page 2")

        doc.save(ARTIFACTS_DIR + "WorkingWithHeadersAndFooters.different_first_page.docx")
        #ExEnd:DifferentFirstPage

    def test_odd_even_pages(self):
        #ExStart:OddEvenPages
        #GistId:2e1b2b28253780881d116e3a873ee668
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.page_setup.odd_and_even_pages_header_footer = True
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_EVEN)
        builder.write("Header for even pages.")
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.write("Header for odd pages.")
        builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_EVEN)
        builder.write("Footer for even pages.")
        builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)
        builder.write("Footer for odd pages.")

        builder.move_to_section(0)
        builder.writeln("Page 1")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln("Page 2")

        doc.save(ARTIFACTS_DIR + "WorkingWithHeadersAndFooters.odd_even_pages.docx")
        #ExEnd:OddEvenPages

    def test_insert_image(self):
        #ExStart:InsertImage
        #GistId:2e1b2b28253780881d116e3a873ee668
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.insert_image(IMAGES_DIR + "Logo.jpg",
                             aw.drawing.RelativeHorizontalPosition.RIGHT_MARGIN,
                             10, aw.drawing.RelativeVerticalPosition.PAGE, 10, 50, 50,
                             aw.drawing.WrapType.THROUGH)

        doc.save(ARTIFACTS_DIR + "WorkingWithHeadersAndFooters.insert_image.docx")
        #ExEnd:InsertImage

    def test_font_props(self):
        #ExStart:FontProps
        #GistId:2e1b2b28253780881d116e3a873ee668
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
        builder.font.name = "Arial"
        builder.font.bold = True
        builder.font.size = 14
        builder.write("Header for pages.")

        doc.save(ARTIFACTS_DIR + "WorkingWithHeadersAndFooters.font_props.docx")
        #ExEnd:FontProps

    def test_page_numbers(self):
        #ExStart:PageNumbers
        #GistId:2e1b2b28253780881d116e3a873ee668
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
        builder.write("Page ")
        builder.insert_field("PAGE", "")
        builder.write(" of ")
        builder.insert_field("NUMPAGES", "")

        doc.save(ARTIFACTS_DIR + "WorkingWithHeadersAndFooters.page_numbers.docx")
        #ExEnd:PageNumbers

    def test_link_to_previous_header_footer(self):
        #ExStart:LinkToPreviousHeaderFooter
        #GistId:2e1b2b28253780881d116e3a873ee668
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.page_setup.different_first_page_header_footer = True

        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_FIRST)
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
        builder.font.name = "Arial"
        builder.font.bold = True
        builder.font.size = 14
        builder.write("Header for the first page.")

        builder.move_to_document_end()
        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)

        current_section = builder.current_section
        page_setup = current_section.page_setup
        page_setup.orientation = aw.Orientation.LANDSCAPE
        # This section does not need a different first-page header/footer we need only one title page in the document,
        # and the header/footer for this page has already been defined in the previous section.
        page_setup.different_first_page_header_footer = False

        # This section displays headers/footers from the previous section
        # by default call currentSection.headers_footers.link_to_previous(false) to cancel this page width
        # is different for the new section, and therefore we need to set different cell widths for a footer table.
        current_section.headers_footers.link_to_previous(False)
        current_section.headers_footers.clear()

        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
        builder.font.name = "Arial"
        builder.font.size = 12
        builder.write("New Header for the first page.")

        doc.save(ARTIFACTS_DIR + "WorkingWithHeadersAndFooters.link_to_previous_header_footer.docx")
        #ExEnd:LinkToPreviousHeaderFooter

    #ExStart:CopyHeadersFootersFromPreviousSection
    #GistId:2e1b2b28253780881d116e3a873ee668
    @staticmethod
    def copy_headers_footers_from_previous_section(section):
        """Clones and copies headers/footers form the previous section to the specified section."""

        previous_section = section.previous_sibling.as_section()
        if previous_section is None:
            return

        section.headers_footers.clear()

        for header_footer in previous_section.headers_footers:
            section.headers_footers.add(header_footer.clone(True))
    #ExEnd:CopyHeadersFootersFromPreviousSection
