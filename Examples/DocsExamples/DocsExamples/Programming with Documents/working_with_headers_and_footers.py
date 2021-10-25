import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

from docs_examples_base import DocsExamplesBase, ARTIFACTS_DIR, IMAGES_DIR

import aspose.words as aw

class WorkingWithHeadersAndFooters(DocsExamplesBase):

    def test_create_header_footer(self):

        #ExStart:CreateHeaderFooterUsingDocBuilder
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        current_section = builder.current_section
        page_setup = current_section.page_setup
        # Specify if we want headers/footers of the first page to be different from other pages.
        # You can also use PageSetup.odd_and_even_pages_header_footer property to specify
        # different headers/footers for odd and even pages.
        page_setup.different_first_page_header_footer = True
        page_setup.header_distance = 20

        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_FIRST)
        builder.paragraph_format.alignment = aw.ParagraphAlignment.CENTER

        builder.font.name = "Arial"
        builder.font.bold = True
        builder.font.size = 14

        builder.write("Aspose.words Header/Footer Creation Primer - Title Page.")

        page_setup.header_distance = 20
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)

        # Insert a positioned image into the top/left corner of the header.
        # Distance from the top/left edges of the page is set to 10 points.
        builder.insert_image(IMAGES_DIR + "Graphics Interchange Format.gif", aw.drawing.RelativeHorizontalPosition.PAGE, 10,
            aw.drawing.RelativeVerticalPosition.PAGE, 10, 50, 50, aw.drawing.WrapType.THROUGH)

        builder.paragraph_format.alignment = aw.ParagraphAlignment.RIGHT

        builder.write("Aspose.words Header/Footer Creation Primer.")

        builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)

        # We use a table with two cells to make one part of the text on the line (with page numbering).
        # To be aligned left, and the other part of the text (with copyright) to be aligned right.
        builder.start_table()

        builder.cell_format.clear_formatting()

        builder.insert_cell()

        builder.cell_format.preferred_width = aw.tables.PreferredWidth.from_percent(100 / 3)

        # It uses PAGE and NUMPAGES fields to auto calculate the current page number and many pages.
        builder.write("Page ")
        builder.insert_field("PAGE", "")
        builder.write(" of ")
        builder.insert_field("NUMPAGES", "")

        builder.current_paragraph.paragraph_format.alignment = aw.ParagraphAlignment.LEFT

        builder.insert_cell()

        builder.cell_format.preferred_width = aw.tables.PreferredWidth.from_percent(100 * 2 / 3)

        builder.write("(C) 2001 Aspose Pty Ltd. All rights reserved.")

        builder.current_paragraph.paragraph_format.alignment = aw.ParagraphAlignment.RIGHT

        builder.end_row()
        builder.end_table()

        builder.move_to_document_end()

        # Make a page break to create a second page on which the primary headers/footers will be seen.
        builder.insert_break(aw.BreakType.PAGE_BREAK)
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

        # If we want to use the already existing header/footer set for this section.
        # But with some minor modifications, then it may be expedient to copy headers/footers
        # from the previous section and apply the necessary modifications where we want them.
        self.copy_headers_footers_from_previous_section(current_section)

        primary_footer = current_section.headers_footers[2] #aw.HeaderFooterType.FOOTER_PRIMARY

        row = primary_footer.tables[0].first_row
        row.first_cell.cell_format.preferred_width = aw.tables.PreferredWidth.from_percent(100 / 3)
        row.last_cell.cell_format.preferred_width = aw.tables.PreferredWidth.from_percent(100 * 2 / 3)

        doc.save(ARTIFACTS_DIR + "WorkingWithHeadersAndFooters.create_header_footer.docx")
        #ExEnd:CreateHeaderFooterUsingDocBuilder


    #ExStart:CopyHeadersFootersFromPreviousSection
    # <summary>
    # Clones and copies headers/footers form the previous section to the specified section.
    # </summary>
    @staticmethod
    def copy_headers_footers_from_previous_section(section):

        previous_section = section.previous_sibling.as_section()

        if previous_section is None:
            return

        section.headers_footers.clear()

        for header_footer in previous_section.headers_footers:
            section.headers_footers.add(header_footer.clone(True))

    #ExEnd:CopyHeadersFootersFromPreviousSection


if __name__ == '__main__':
    unittest.main()
