import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class WorkingWithTxtSaveOptions(DocsExamplesBase):

    def test_add_bidi_marks(self):

        #ExStart:AddBidiMarks
        #GistId:5e05b790a4d2258054abbf842ce6e427
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello world!")
        builder.paragraph_format.bidi = True
        builder.writeln("שלום עולם!")
        builder.writeln("مرحبا بالعالم!")

        save_options = aw.saving.TxtSaveOptions()
        save_options.add_bidi_marks = True

        doc.save(ARTIFACTS_DIR + "WorkingWithTxtSaveOptions.add_bidi_marks.txt", save_options)
        #ExEnd:AddBidiMarks

    def test_use_tab_for_list_indentation(self):

        #ExStart:UseTabForListIndentation
        #GistId:5e05b790a4d2258054abbf842ce6e427
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create a list with three levels of indentation.
        builder.list_format.apply_number_default()
        builder.writeln("Item 1")
        builder.list_format.list_indent()
        builder.writeln("Item 2")
        builder.list_format.list_indent()
        builder.write("Item 3")

        save_options = aw.saving.TxtSaveOptions()
        save_options.list_indentation.count = 1
        save_options.list_indentation.character = '\t'

        doc.save(ARTIFACTS_DIR + "WorkingWithTxtSaveOptions.use_tab_for_list_indentation.txt", save_options)
        #ExEnd:UseTabForListIndentation

    def test_use_space_for_list_indentation(self):

        #ExStart:UseSpaceForListIndentation
        #GistId:5e05b790a4d2258054abbf842ce6e427
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Create a list with three levels of indentation.
        builder.list_format.apply_number_default()
        builder.writeln("Item 1")
        builder.list_format.list_indent()
        builder.writeln("Item 2")
        builder.list_format.list_indent()
        builder.write("Item 3")

        save_options = aw.saving.TxtSaveOptions()
        save_options.list_indentation.count = 3
        save_options.list_indentation.character = ' '

        doc.save(ARTIFACTS_DIR + "WorkingWithTxtSaveOptions.use_space_for_list_indentation.txt", save_options)
        #ExEnd:UseSpaceForListIndentation

    def test_export_headers_footers_mode(self):

        #ExStart:ExportHeadersFootersMode
        #GistId:5e05b790a4d2258054abbf842ce6e427
        doc = aw.Document()

        # Insert even and primary headers/footers into the document.
        # The primary header/footers will override the even headers/footers.
        doc.first_section.headers_footers.add(aw.HeaderFooter(doc, aw.HeaderFooterType.HEADER_EVEN))
        doc.first_section.headers_footers.get_by_header_footer_type(aw.HeaderFooterType.HEADER_EVEN).append_paragraph("Even header")
        doc.first_section.headers_footers.add(aw.HeaderFooter(doc, aw.HeaderFooterType.FOOTER_EVEN))
        doc.first_section.headers_footers.get_by_header_footer_type(aw.HeaderFooterType.FOOTER_EVEN).append_paragraph("Even footer")
        doc.first_section.headers_footers.add(aw.HeaderFooter(doc, aw.HeaderFooterType.HEADER_PRIMARY))
        doc.first_section.headers_footers.get_by_header_footer_type(aw.HeaderFooterType.HEADER_PRIMARY).append_paragraph("Primary header")
        doc.first_section.headers_footers.add(aw.HeaderFooter(doc, aw.HeaderFooterType.FOOTER_PRIMARY))
        doc.first_section.headers_footers.get_by_header_footer_type(aw.HeaderFooterType.FOOTER_PRIMARY).append_paragraph("Primary footer")

        # Insert pages to display these headers and footers.
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Page 1")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln("Page 2")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.write("Page 3")

        options = aw.saving.TxtSaveOptions()
        options.save_format = aw.SaveFormat.TEXT

        # All headers and footers are placed at the very end of the output document.
        options.export_headers_footers_mode = aw.saving.TxtExportHeadersFootersMode.ALL_AT_END
        doc.save(ARTIFACTS_DIR + "WorkingWithTxtSaveOptions.HeadersFootersMode.AllAtEnd.txt", options)

        # Only primary headers and footers are exported at the beginning and end of each section.
        options.export_headers_footers_mode = aw.saving.TxtExportHeadersFootersMode.PRIMARY_ONLY
        doc.save(ARTIFACTS_DIR + "WorkingWithTxtSaveOptions.HeadersFootersMode.PrimaryOnly.txt", options)

        # No headers and footers are exported.
        options.export_headers_footers_mode = aw.saving.TxtExportHeadersFootersMode.NONE
        doc.save(ARTIFACTS_DIR + "WorkingWithTxtSaveOptions.HeadersFootersMode.None.txt", options)
        #ExEnd:ExportHeadersFootersMode
