# Copyright (c) 2001-2022 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR

#class ExPrinting(ApiExampleBase):

#    # Run only when the printer driver is installed.
#    def test_custom_print(self):

#        #ExStart
#        #ExFor:PageInfo.get_dot_net_paper_size
#        #ExFor:PageInfo.landscape
#        #ExSummary:Shows how to customize the printing of Aspose.Words documents.
#        doc = aw.Document(MY_DIR + "Rendering.docx")

#        print_doc = ExPrinting.MyPrintDocument(doc)
#        print_doc.printer_settings.print_range = drawing.printing.PrintRange.SOME_PAGES
#        print_doc.printer_settings.from_page = 1
#        print_doc.printer_settings.to_page = 1

#        print_doc.print()

#    class MyPrintDocument(drawing.printing.PrintDocument):
#        """Selects an appropriate paper size, orientation, and paper tray when printing."""

#        def __init__(self, document: aw.Document):
#            self.document = document
#            self.current_page = 0
#            self.page_to = 0

#        def on_begin_print(self, e: drawing.printing.PrintEventArgs):
#            """Initializes the range of pages to be printed according to the user selection."""

#            base.on_begin_print(e)

#            if self.printer_settings.print_range == drawing.printing.PrintRange.ALL_PAGES:
#                self.current_page = 1
#                self.page_to = self.document.page_count
#            elif self.printer_settings.print_range == drawing.printing.PrintRange.SOME_PAGES:
#                self.current_page = drawing.printing.PrinterSettings.FROM_PAGE
#                self.page_to = drawing.printing.PrinterSettings.TO_PAGE
#            else:
#                raise Exception("Unsupported print range.")

#        def on_query_page_settings(self, e: drawing.printing.QueryPageSettingsEventArgs):
#            """Called before each page is printed."""

#            base.on_query_page_settings(e)

#            # A single Microsoft Word document can have multiple sections that specify pages with different sizes,
#            # orientations, and paper trays. The .NET printing framework calls this code before
#            # each page is printed, which gives us a chance to specify how to print the current page.
#            page_info = self.document.get_page_info(self.current_page - 1)
#            e.page_settings.paper_size = page_info.get_dot_net_paper_size(drawing.printing.PrinterSettings.PAPER_SIZES)

#            # Microsoft Word stores the paper source (printer tray) for each section as a printer-specific value.
#            # To obtain the correct tray value, you will need to use the "raw_kind" property, which your printer should return.
#            e.page_settings.paper_source.raw_kind = page_info.paper_tray
#            e.page_settings.landscape = page_info.landscape

#        def on_print_page(self, e: drawing.printing.PrintPageEventArgs):
#            """Called for each page to render it for printing."""

#            base.on_print_page(e)

#            # Aspose.Words rendering engine creates a page drawn from the origin (x = 0, y = 0) of the paper.
#            # There will be a hard margin in the printer, which will render each page. We need to offset by that hard margin.
#            hard_offset_x = 0.0
#            hard_offset_y = 0.0

#            # Below are two ways of setting a hard margin.
#            if e.page_settings is not None and e.page_settings.hard_margin_x != 0 and e.page_settings.hard_margin_y != 0:
#                # 1 -  Via the "page_settings" property.
#                hard_offset_x = e.page_settings.hard_margin_x
#                hard_offset_y = e.page_settings.hard_margin_y
#            else:
#                # 2 -  Using our own values, if the "PageSettings" property is unavailable.
#                hard_offset_x = 20
#                hard_offset_y = 20

#            self.document.render_to_scale(self.currentPage, e.graphics, -hard_offset_x, -hard_offset_y, 1.0)

#            self.current_page += 1
#            e.has_more_pages = self.current_page <= self.page_to

#    #ExEnd

#    # Run only when the printer driver is installed.
#    def test_print_page_info(self):

#        #ExStart
#        #ExFor:PageInfo
#        #ExFor:PageInfo.get_size_in_pixels(float,float,float)
#        #ExFor:PageInfo.get_specified_printer_paper_source(PaperSourceCollection,PaperSource)
#        #ExFor:PageInfo.height_in_points
#        #ExFor:PageInfo.landscape
#        #ExFor:PageInfo.paper_size
#        #ExFor:PageInfo.paper_tray
#        #ExFor:PageInfo.size_in_points
#        #ExFor:PageInfo.width_in_points
#        #ExSummary:Shows how to print page size and orientation information for every page in a Word document.
#        doc = aw.Document(MY_DIR + "Rendering.docx")

#        # The first section has 2 pages. We will assign a different printer paper tray to each one,
#        # whose number will match a kind of paper source. These sources and their Kinds will vary
#        # depending on the installed printer driver.
#        paper_sources = drawing.printing.PrinterSettings().paper_sources

#        doc.first_section.page_setup.first_page_tray = paper_sources[0].raw_kind
#        doc.first_section.page_setup.other_pages_tray = paper_sources[1].raw_kind

#        print("Document \"{0}\" contains {1} pages.", doc.original_file_name, doc.page_count)

#        scale = 1.0
#        dpi = 96.0

#        for i in range(doc.page_count):

#            # Each page has a PageInfo object, whose index is the respective page's number.
#            page_info = doc.get_page_info(i)

#            # Print the page's orientation and dimensions.
#            print(f"Page {i + 1}:")
#            print(f"\tOrientation:\t{'Landscape' if page_info.landscape else 'Portrait'}")
#            print(f"\tPaper size:\t\t{page_info.paper_size} ({page_info.width_in_points:F0}x{page_info.height_in_points:F0}pt)")
#            print(f"\tSize in points:\t{page_info.size_in_points}")
#            print(f"\tSize in pixels:\t{page_info.get_size_in_pixels(1.0, 96)} at {scale * 100}% scale, {dpi} dpi")

#            # Print the source tray information.
#            print(f"\tTray:\t{page_info.paper_tray}")
#            source = page_info.get_specified_printer_paper_source(paper_sources, paper_sources[0])
#            print(f"\tSuitable print source:\t{source.source_name}, kind: {source.kind}")

#        #ExEnd

#    # Run only when the printer driver is installed.
#    def test_printer_settings_container(self):

#        #ExStart
#        #ExFor:PrinterSettingsContainer
#        #ExFor:PrinterSettingsContainer.__init__(PrinterSettings)
#        #ExFor:PrinterSettingsContainer.default_page_settings_paper_source
#        #ExFor:PrinterSettingsContainer.paper_sizes
#        #ExFor:PrinterSettingsContainer.paper_sources
#        #ExSummary:Shows how to access and list your printer's paper sources and sizes.
#        # The "PrinterSettingsContainer" contains a "PrinterSettings" object,
#        # which contains unique data for different printer drivers.
#        container = aw.rendering.PrinterSettingsContainer(PrinterSettings())

#        print(f"This printer contains {container.paper_sources.count} printer paper sources:")
#        for paper_source in container.paper_sources:

#            is_default = container.default_page_settings_paper_source.source_name == paper_source.source_name
#            print(f"\t{paperSource.SourceName}, RawKind: {paperSource.RawKind} {'(Default)' if is_default else ''}")

#        # The "paper_sizes" property contains the list of paper sizes to instruct the printer to use.
#        # Both the PrinterSource and PrinterSize contain a "raw_kind" property,
#        # which equates to a paper type listed on the PaperSourceKind enum.
#        # If there is a paper source with the same "raw_kind" value as that of the printing page,
#        # the printer will print the page using the provided paper source and size.
#        # Otherwise, the printer will default to the source designated by the "default_page_settings_paper_source" property.
#        print(f"{container.paper_sizes.count} paper sizes:")
#        for paper_size in container.paper_sizes:

#            print(f"\t{paper_size}, RawKind: {paper_size.raw_kind}")

#        #ExEnd

#    # Run only when the printer driver is installed.
#    def test_print(self):

#        #ExStart
#        #ExFor:Document.print
#        #ExFor:Document.print(str)
#        #ExSummary:Shows how to print a document using the default printer.
#        doc = aw.Document()
#        builder = aw.DocumentBuilder(doc)
#        builder.writeln("Hello world!")

#        # Below are two ways of printing our document.
#        # 1 -  Print using the default printer:
#        doc.print()

#        # 2 -  Specify a printer that we wish to print the document with by name:
#        my_printer = drawing.settings.PrinterSettings.installed_printers[4]

#        self.assertEqual("HPDAAB96 (HP ENVY 5000 series)", my_printer)

#        doc.print(my_printer)
#        #ExEnd

#    # Run only when the printer driver is installed.
#    def test_print_range(self):

#        #ExStart
#        #ExFor:Document.print(PrinterSettings)
#        #ExFor:Document.print(PrinterSettings,str)
#        #ExSummary:Shows how to print a range of pages.
#        doc = aw.Document(MY_DIR + "Rendering.docx")

#        # Create a "PrinterSettings" object to modify how we print the document.
#        printer_settings = drawing.printing.PrinterSettings()

#        # Set the "print_range" property to "PrintRange.SOME_PAGES" to
#        # tell the printer that we intend to print only some document pages.
#        printer_settings.print_range = drawing.printing.PrintRange.SOME_PAGES

#        # Set the "from_page" property to "1", and the "to_page" property to "3" to print pages 1 through to 3.
#        # Page indexing is 1-based.
#        printer_settings.from_page = 1
#        printer_settings.to_page = 3

#        # Below are two ways of printing our document.
#        # 1 -  Print while applying our printing settings:
#        doc.print(printer_settings)

#        # 2 -  Print while applying our printing settings, while also
#        # giving the document a custom name that we may recognize in the printer queue:
#        doc.print(printer_settings, "My rendered document")
#        #ExEnd

#    # Run only when the printer driver is installed.
#    def test_preview_and_print(self):

#        #ExStart
#        #ExFor:AsposeWordsPrintDocument.__init__(Document)
#        #ExFor:AsposeWordsPrintDocument.cache_printer_settings
#        #ExSummary:Shows how to select a page range and a printer to print the document with, and then bring up a print preview.
#        doc = aw.Document(MY_DIR + "Rendering.docx")

#        preview_dlg = PrintPreviewDialog()

#        # Call the "show" method to get the print preview form to show on top.
#        preview_dlg.show()

#        # Initialize the Print Dialog with the number of pages in the document.
#        print_dlg = PrintDialog()
#        print_dlg.allow_some_pages = True
#        print_dlg.printer_settings.minimum_page = 1
#        print_dlg.printer_settings.maximum_page = doc.page_count
#        print_dlg.printer_settings.from_page = 1
#        print_dlg.printer_settings.to_page = doc.page_count

#        if print_dlg.show_dialog() != DialogResult.OK:
#            return

#        # Create the "Aspose.Words" implementation of the .NET print document,
#        # and then pass the printer settings from the dialog.
#        aw_print_doc = aw.rendering.AsposeWordsPrintDocument(doc)
#        aw_print_doc.printer_settings = print_dlg.printer_settings

#        # Use the "cache_printer_settings" method to reduce time of the first call of the "Print" method.
#        aw_print_doc.cache_printer_settings()

#        # Call the "hide", and then the "invalidate_preview" methods to get the print preview to show on top.
#        preview_dlg.hide()
#        preview_dlg.print_preview_control.invalidate_preview()

#        # Pass the "Aspose.Words" print document to the .NET Print Preview dialog.
#        preview_dlg.document = aw_print_doc

#        preview_dlg.show_dialog()
#        #ExEnd
