import unittest
import io

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

class ExViewOptions(ApiExampleBase):

    def test_set_zoom_percentage(self):

        #ExStart
        #ExFor:Document.view_options
        #ExFor:ViewOptions
        #ExFor:ViewOptions.view_type
        #ExFor:ViewOptions.zoom_percent
        #ExFor:ViewOptions.zoom_type
        #ExFor:ViewType
        #ExSummary:Shows how to set a custom zoom factor, which older versions of Microsoft Word will apply to a document upon loading.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello world!")

        doc.view_options.view_type = aw.settings.ViewType.PAGE_LAYOUT
        doc.view_options.zoom_percent = 50

        self.assertEqual(aw.settings.ZoomType.CUSTOM, doc.view_options.zoom_type)
        self.assertEqual(aw.settings.ZoomType.NONE, doc.view_options.zoom_type)

        doc.save(ARTIFACTS_DIR + "ViewOptions.SetZoomPercentage.doc")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "ViewOptions.SetZoomPercentage.doc")

        self.assertEqual(aw.settings.ViewType.PAGE_LAYOUT, doc.view_options.view_type)
        self.assertEqual(50.0, doc.view_options.zoom_percent)
        self.assertEqual(aw.settings.ZoomType.NONE, doc.view_options.zoom_type)

    def test_set_zoom_type(self):

        for zoom_type in (aw.settings.ZoomType.PAGE_WIDTH, aw.settings.ZoomType.FULL_PAGE, aw.settings.ZoomType.TEXT_FIT):
            with self.subTest(zoom_type=zoom_type):
                #ExStart
                #ExFor:Document.view_options
                #ExFor:ViewOptions
                #ExFor:ViewOptions.zoom_type
                #ExSummary:Shows how to set a custom zoom type, which older versions of Microsoft Word will apply to a document upon loading.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.writeln("Hello world!")

                # Set the "zoom_type" property to "ZoomType.PAGE_WIDTH" to get Microsoft Word
                # to automatically zoom the document to fit the width of the page.
                # Set the "zoom_type" property to "ZoomType.FULL_PAGE" to get Microsoft Word
                # to automatically zoom the document to make the entire first page visible.
                # Set the "zoom_type" property to "ZoomType.TEXT_FIT" to get Microsoft Word
                # to automatically zoom the document to fit the inner text margins of the first page.
                doc.view_options.zoom_type = zoom_type

                doc.save(ARTIFACTS_DIR + "ViewOptions.SetZoomType.doc")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "ViewOptions.SetZoomType.doc")

                self.assertEqual(zoom_type, doc.view_options.zoom_type)

    def test_display_background_shape(self):

        for display_background_shape in (False, True):
            with self.subTest(display_background_shape=display_background_shape):
                #ExStart
                #ExFor:ViewOptions.display_background_shape
                #ExSummary:Shows how to hide/display document background images in view options.
                # Use an HTML string to create a new document with a flat background color.
                html = """
                    <html>
                        <body style='background-color: blue'>
                            <p>Hello world!</p>
                        </body>
                    </html>"""

                doc = aw.Document(io.BytesIO(html.encode('utf-8')))

                # The source for the document has a flat color background,
                # the presence of which will set the "display_background_shape" flag to "True".
                self.assertTrue(doc.view_options.display_background_shape)

                # Keep the "display_background_shape" as "True" to get the document to display the background color.
                # This may affect some text colors to improve visibility.
                # Set the "display_background_shape" to "False" to not display the background color.
                doc.view_options.display_background_shape = display_background_shape

                doc.save(ARTIFACTS_DIR + "ViewOptions.DisplayBackgroundShape.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "ViewOptions.DisplayBackgroundShape.docx")

                self.assertEqual(display_background_shape, doc.view_options.display_background_shape)

    def test_display_page_boundaries(self):
        
        for do_not_display_page_boundaries in (False, True):
            with self.subTest(do_not_display_page_boundaries=do_not_display_page_boundaries):
                #ExStart
                #ExFor:ViewOptions.do_not_display_page_boundaries
                #ExSummary:Shows how to hide vertical whitespace and headers/footers in view options.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Insert content that spans across 3 pages.
                builder.writeln("Paragraph 1, Page 1.")
                builder.insert_break(aw.BreakType.PAGE_BREAK)
                builder.writeln("Paragraph 2, Page 2.")
                builder.insert_break(aw.BreakType.PAGE_BREAK)
                builder.writeln("Paragraph 3, Page 3.")

                # Insert a header and a footer.
                builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
                builder.writeln("This is the header.")
                builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)
                builder.writeln("This is the footer.")

                # This document contains a small amount of content that takes up a few full pages worth of space.
                # Set the "do_not_display_page_boundaries" flag to "True" to get older versions of Microsoft Word to omit headers,
                # footers, and much of the vertical whitespace when displaying our document.
                # Set the "do_not_display_page_boundaries" flag to "False" to get older versions of Microsoft Word
                # to normally display our document.
                doc.view_options.do_not_display_page_boundaries = do_not_display_page_boundaries

                doc.save(ARTIFACTS_DIR + "ViewOptions.DisplayPageBoundaries.doc")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "ViewOptions.DisplayPageBoundaries.doc")

                self.assertEqual(do_not_display_page_boundaries, doc.view_options.do_not_display_page_boundaries)

    def test_forms_design(self):

        for use_forms_design in (False, True):
            with self.subTest(use_forms_design=use_forms_design):
                #ExStart
                #ExFor:ViewOptions.forms_design
                #ExSummary:Shows how to enable/disable forms design mode.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.writeln("Hello world!")

                # Set the "forms_design" property to "False" to keep forms design mode disabled.
                # Set the "forms_design" property to "True" to enable forms design mode.
                doc.view_options.forms_design = use_forms_design

                doc.save(ARTIFACTS_DIR + "ViewOptions.FormsDesign.xml")

                with open(ARTIFACTS_DIR + "ViewOptions.FormsDesign.xml") as file:
                    self.assertEqual(use_forms_design, "<w:formsDesign />" in file.read())
                #ExEnd
