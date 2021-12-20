# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

from datetime import date

import aspose.words as aw

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, IMAGE_DIR

class ExHeaderFooter(ApiExampleBase):

    def test_create(self):

        #ExStart
        #ExFor:HeaderFooter
        #ExFor:HeaderFooter.__init__(DocumentBase,HeaderFooterType)
        #ExFor:HeaderFooter.header_footer_type
        #ExFor:HeaderFooter.is_header
        #ExFor:HeaderFooterCollection
        #ExFor:Paragraph.is_end_of_header_footer
        #ExFor:Paragraph.parent_section
        #ExFor:Paragraph.parent_story
        #ExFor:Story.append_paragraph
        #ExSummary:Shows how to create a header and a footer.
        doc = aw.Document()

        # Create a header and append a paragraph to it. The text in that paragraph
        # will appear at the top of every page of this section, above the main body text.
        header = aw.HeaderFooter(doc, aw.HeaderFooterType.HEADER_PRIMARY)
        doc.first_section.headers_footers.add(header)

        para = header.append_paragraph("My header.")

        self.assertTrue(header.is_header)
        self.assertTrue(para.is_end_of_header_footer)

        # Create a footer and append a paragraph to it. The text in that paragraph
        # will appear at the bottom of every page of this section, below the main body text.
        footer = aw.HeaderFooter(doc, aw.HeaderFooterType.FOOTER_PRIMARY)
        doc.first_section.headers_footers.add(footer)

        para = footer.append_paragraph("My footer.")

        self.assertFalse(footer.is_header)
        self.assertTrue(para.is_end_of_header_footer)

        self.assertEqual(footer, para.parent_story)
        self.assertEqual(footer.parent_section, para.parent_section)
        self.assertEqual(footer.parent_section, header.parent_section)

        doc.save(ARTIFACTS_DIR + "HeaderFooter.create.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "HeaderFooter.create.docx")

        self.assertIn("My header.", doc.first_section.headers_footers[aw.HeaderFooterType.HEADER_PRIMARY].range.text)
        self.assertIn("My footer.", doc.first_section.headers_footers[aw.HeaderFooterType.FOOTER_PRIMARY].range.text)

    def test_link(self):

        #ExStart
        #ExFor:HeaderFooter.is_linked_to_previous
        #ExFor:HeaderFooterCollection.__getitem__(int)
        #ExFor:HeaderFooterCollection.link_to_previous(HeaderFooterType,bool)
        #ExFor:HeaderFooterCollection.link_to_previous(bool)
        #ExFor:HeaderFooter.parent_section
        #ExSummary:Shows how to link headers and footers between sections.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Section 1")
        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
        builder.write("Section 2")
        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
        builder.write("Section 3")

        # Move to the first section and create a header and a footer. By default,
        # the header and the footer will only appear on pages in the section that contains them.
        builder.move_to_section(0)

        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.write("This is the header, which will be displayed in sections 1 and 2.")

        builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)
        builder.write("This is the footer, which will be displayed in sections 1, 2 and 3.")

        # We can link a section's headers/footers to the previous section's headers/footers
        # to allow the linking section to display the linked section's headers/footers.
        doc.sections[1].headers_footers.link_to_previous(True)

        # Each section will still have its own header/footer objects. When we link sections,
        # the linking section will display the linked section's header/footers while keeping its own.
        self.assertNotEqual(doc.sections[0].headers_footers[0], doc.sections[1].headers_footers[0])
        self.assertNotEqual(doc.sections[0].headers_footers[0].parent_section, doc.sections[1].headers_footers[0].parent_section)

        # Link the headers/footers of the third section to the headers/footers of the second section.
        # The second section already links to the first section's header/footers,
        # so linking to the second section will create a link chain.
        # The first, second, and now the third sections will all display the first section's headers.
        doc.sections[2].headers_footers.link_to_previous(True)

        # We can un-link a previous section's header/footers by passing "False" when calling the LinkToPrevious method.
        doc.sections[2].headers_footers.link_to_previous(False)

        # We can also select only a specific type of header/footer to link using this method.
        # The third section now will have the same footer as the second and first sections, but not the header.
        doc.sections[2].headers_footers.link_to_previous(aw.HeaderFooterType.FOOTER_PRIMARY, True)

        # The first section's header/footers cannot link themselves to anything because there is no previous section.
        self.assertEqual(2, doc.sections[0].headers_footers.count)
        self.assertEqual(2, len([node for node in doc.sections[0].headers_footers if not node.as_header_footer().is_linked_to_previous]))

        # All the second section's header/footers are linked to the first section's headers/footers.
        self.assertEqual(6, doc.sections[1].headers_footers.count)
        self.assertEqual(6, len([node for node in doc.sections[1].headers_footers if node.as_header_footer().is_linked_to_previous]))

        # In the third section, only the footer is linked to the first section's footer via the second section.
        self.assertEqual(6, doc.sections[2].headers_footers.count)
        self.assertEqual(5, len([node for node in doc.sections[2].headers_footers if not node.as_header_footer().is_linked_to_previous]))
        self.assertTrue(doc.sections[2].headers_footers[3].is_linked_to_previous)

        doc.save(ARTIFACTS_DIR + "HeaderFooter.link.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "HeaderFooter.link.docx")

        self.assertEqual(2, doc.sections[0].headers_footers.count)
        self.assertEqual(2, len([node for node in doc.sections[0].headers_footers if not node.as_header_footer().is_linked_to_previous]))

        self.assertEqual(0, doc.sections[1].headers_footers.count)
        self.assertEqual(0, len([node for node in doc.sections[1].headers_footers if node.as_header_footer().is_linked_to_previous]))

        self.assertEqual(5, doc.sections[2].headers_footers.count)
        self.assertEqual(5, len([node for node in doc.sections[2].headers_footers if not node.as_header_footer().is_linked_to_previous]))

    def test_remove_footers(self):

        #ExStart
        #ExFor:Section.headers_footers
        #ExFor:HeaderFooterCollection
        #ExFor:HeaderFooterCollection.__getitem__(HeaderFooterType)
        #ExFor:HeaderFooter
        #ExSummary:Shows how to delete all footers from a document.
        doc = aw.Document(MY_DIR + "Header and footer types.docx")

        # Iterate through each section and remove footers of every kind.
        for section in doc:
            section = section.as_section()

            # There are three kinds of footer and header types.
            # 1 -  The "First" header/footer, which only appears on the first page of a section.
            footer = section.headers_footers[aw.HeaderFooterType.FOOTER_FIRST]
            if footer is not None:
                footer.remove()

            # 2 -  The "Primary" header/footer, which appears on odd pages.
            footer = section.headers_footers[aw.HeaderFooterType.FOOTER_PRIMARY]
            if footer is not None:
                footer.remove()

            # 3 -  The "Even" header/footer, which appears on odd even pages.
            footer = section.headers_footers[aw.HeaderFooterType.FOOTER_EVEN]
            if footer is not None:
                footer.remove()

            self.assertEqual(0, len([node for node in section.headers_footers if not node.as_header_footer().is_header]))

        doc.save(ARTIFACTS_DIR + "HeaderFooter.remove_footers.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "HeaderFooter.remove_footers.docx")

        self.assertEqual(1, doc.sections.count)
        self.assertEqual(0, len([node for node in doc.first_section.headers_footers if not node.as_header_footer().is_header]))
        self.assertEqual(3, len([node for node in doc.first_section.headers_footers if node.as_header_footer().is_header]))

    def test_export_mode(self):

        #ExStart
        #ExFor:HtmlSaveOptions.export_headers_footers_mode
        #ExFor:ExportHeadersFootersMode
        #ExSummary:Shows how to omit headers/footers when saving a document to HTML.
        doc = aw.Document(MY_DIR + "Header and footer types.docx")

        # This document contains headers and footers. We can access them via the "headers_footers" collection.
        self.assertEqual("First header", doc.first_section.headers_footers[aw.HeaderFooterType.HEADER_FIRST].get_text().strip())

        # Formats such as .html do not split the document into pages, so headers/footers will not function the same way
        # they would when we open the document as a .docx using Microsoft Word.
        # If we convert a document with headers/footers to html, the conversion will assimilate the headers/footers into body text.
        # We can use a SaveOptions object to omit headers/footers while converting to html.
        save_options = aw.saving.HtmlSaveOptions(aw.SaveFormat.HTML)
        save_options.export_headers_footers_mode = aw.saving.ExportHeadersFootersMode.NONE

        doc.save(ARTIFACTS_DIR + "HeaderFooter.export_mode.html", save_options)

        # Open our saved document and verify that it does not contain the header's text
        doc = aw.Document(ARTIFACTS_DIR + "HeaderFooter.export_mode.html")

        self.assertNotIn("First header", doc.range.text)
        #ExEnd

    def test_replace_text(self):

        #ExStart
        #ExFor:Document.first_section
        #ExFor:Section.headers_footers
        #ExFor:HeaderFooterCollection.__getitem__(HeaderFooterType)
        #ExFor:HeaderFooter
        #ExFor:Range.replace(str,str,FindReplaceOptions)
        #ExSummary:Shows how to replace text in a document's footer.
        doc = aw.Document(MY_DIR + "Footer.docx")

        headers_footers = doc.first_section.headers_footers
        footer = headers_footers[aw.HeaderFooterType.FOOTER_PRIMARY]

        options = aw.replacing.FindReplaceOptions()
        options.match_case = False
        options.find_whole_words_only = False

        current_year = date.today().year
        footer.range.replace("(C) 2006 Aspose Pty Ltd.", f"Copyright (C) {current_year} by Aspose Pty Ltd.", options)

        doc.save(ARTIFACTS_DIR + "HeaderFooter.replace_text.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "HeaderFooter.replace_text.docx")

        self.assertIn(f"Copyright (C) {currentYear} by Aspose Pty Ltd.", doc.range.text)

    ##ExStart
    ##ExFor:IReplacingCallback
    ##ExFor:PageSetup.different_first_page_header_footer
    ##ExSummary:Shows how to track the order in which a text replacement operation traverses nodes.
    #def test_order(self):

    #    for different_first_page_header_footer in (False, True):
    #        with self.subTest(different_first_page_header_footer=different_first_page_header_footer):
    #            doc = aw.Document(MY_DIR + "Header and footer types.docx")

    #            first_page_section = doc.first_section

    #            logger = ExHeaderFooter.ReplaceLog()
    #            options = aw.replacing.FindReplaceOptions()
    #            options.replacing_callback = logger

    #            # Using a different header/footer for the first page will affect the search order.
    #            first_page_section.page_setup.different_first_page_header_footer = different_first_page_header_footer
    #            doc.range.replace_regex("(header|footer)", "", options)

    #            if different_first_page_header_footer:
    #                self.assertEqual("First header\nFirst footer\nSecond header\nSecond footer\nThird header\nThird footer\n",
    #                    logger.text.replace("\r", ""))
    #            else:
    #                self.assertEqual("Third header\nFirst header\nThird footer\nFirst footer\nSecond header\nSecond footer\n",
    #                    logger.text.replace("\r", ""))

    #class ReplaceLog(aw.replacing.IReplacingCallback):
    #    """During a find-and-replace operation, records the contents of every node that has text that the operation 'finds',
    #    in the state it is in before the replacement takes place.
    #    This will display the order in which the text replacement operation traverses nodes."""

    #    def __init__(self):
    #        self.text_builder = io.StringIO()

    #    def replacing(self, args: aw.replacing.ReplacingArgs) -> aw.replacing.ReplaceAction:

    #        self.text_builder.write(args.match_node.get_text() + '\n')
    #        return aw.replacing.ReplaceAction.SKIP

    #    @property
    #    def text(self):
    #        return self.text_builder.getvalue()
    ##ExEnd

    def test_primer(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        current_section = builder.current_section
        page_setup = current_section.page_setup

        # Specify if we want headers/footers of the first page to be different from other pages.
        # You can also use "PageSetup.odd_and_even_pages_header_footer" property to specify
        # different headers/footers for odd and even pages.
        page_setup.different_first_page_header_footer = True

        # Create header for the first page.
        page_setup.header_distance = 20
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_FIRST)
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER

        builder.font.name = "Arial"
        builder.font.bold = True
        builder.font.size = 14
        builder.write("Aspose.Words Header/Footer Creation Primer - Title Page.")

        # Create header for pages other than first.
        page_setup.header_distance = 20
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)

        # Insert an absolutely positioned image into the top/left corner of the header.
        # Distance from the top/left edges of the page is set to 10 points.
        image_file_name = IMAGE_DIR + "Logo.jpg"
        builder.insert_image(image_file_name, aw.drawing.RelativeHorizontalPosition.PAGE, 10, aw.drawing.RelativeVerticalPosition.PAGE, 10,
            50, 50, aw.drawing.WrapType.THROUGH)

        builder.paragraph_format.alignment = aw.ParagraphAlignment.RIGHT
        builder.write("Aspose.Words Header/Footer Creation Primer.")

        # Create footer for pages other than first.
        builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)

        # We use a table with two cells to make one part of the text on the line (with page numbering)
        # to be aligned left, and the other part of the text (with copyright) to be aligned right.
        builder.start_table()

        builder.cell_format.clear_formatting()

        builder.insert_cell()

        builder.cell_format.preferred_width = aw.tables.PreferredWidth.from_percent(100.0 / 3)

        # Insert page numbering text here.
        # It uses PAGE and NUMPAGES fields to auto calculate the current page number and a total number of pages.
        builder.write("Page ")
        builder.insert_field("PAGE", "")
        builder.write(" of ")
        builder.insert_field("NUMPAGES", "")

        builder.current_paragraph.paragraph_format.alignment = aw.ParagraphAlignment.LEFT

        builder.insert_cell()
        builder.cell_format.preferred_width = aw.tables.PreferredWidth.from_percent(100.0 * 2 / 3)

        builder.write("(C) 2001 Aspose Pty Ltd. All rights reserved.")

        builder.current_paragraph.paragraph_format.alignment = aw.ParagraphAlignment.RIGHT

        builder.end_row()
        builder.end_table()

        builder.move_to_document_end()
        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # Make section break to create a third page with a different page orientation.
        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)

        current_section = builder.current_section
        page_setup = current_section.page_setup

        page_setup.orientation = aw.Orientation.LANDSCAPE

        # This section does not need different first page header/footer.
        # We need only one title page in the document and the header/footer for this page
        # has already been defined in the previous section.
        page_setup.different_first_page_header_footer = False

        # This section displays headers/footers from the previous section by default.
        # Call "current_section.headers_footers.link_to_previous(False)" to cancel this.
        # Page width is different for the new section and therefore we need to set
        # a different cell widths for a footer table.
        current_section.headers_footers.link_to_previous(False)

        # If we want to use the already existing header/footer set for this section
        # but with some minor modifications then it may be expedient to copy headers/footers
        # from the previous section and apply the necessary modifications where we want them.
        ExHeaderFooter.copy_headers_footers_from_previous_section(current_section)

        primary_footer = current_section.headers_footers[aw.HeaderFooterType.FOOTER_PRIMARY]

        row = primary_footer.tables[0].first_row
        row.first_cell.cell_format.preferred_width = aw.tables.PreferredWidth.from_percent(100.0 / 3)
        row.last_cell.cell_format.preferred_width = aw.tables.PreferredWidth.from_percent(100.0 * 2 / 3)

        doc.save(ARTIFACTS_DIR + "HeaderFooter.primer.docx")

    @staticmethod
    def copy_headers_footers_from_previous_section(section: aw.Section):
        """Clones and copies headers/footers form the previous section to the specified section."""

        previous_section = section.previous_sibling.as_section()

        if previous_section is None:
            return

        section.headers_footers.clear()

        for header_footer in previous_section.headers_footers:
            header_footer = header_footer.as_header_footer()

            section.headers_footers.add(header_footer.clone(True))
