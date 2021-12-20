# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import os
import glob
import shutil

import aspose.words as aw

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR
from document_helper import DocumentHelper

class ExHtmlFixedSaveOptions(ApiExampleBase):

    def test_use_encoding(self):

        #ExStart
        #ExFor:HtmlFixedSaveOptions.encoding
        #ExSummary:Shows how to set which encoding to use while exporting a document to HTML.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello World!")

        # The default encoding is UTF-8. If we want to represent our document using a different encoding,
        # we can use a SaveOptions object to set a specific encoding.
        html_fixed_save_options = aw.saving.HtmlFixedSaveOptions()
        html_fixed_save_options.encoding = "ascii"

        self.assertEqual("us-ascii", html_fixed_save_options.encoding)

        doc.save(ARTIFACTS_DIR + "HtmlFixedSaveOptions.use_encoding.html", html_fixed_save_options)
        #ExEnd

        with open(ARTIFACTS_DIR + "HtmlFixedSaveOptions.use_encoding.html", "rt", encoding="utf-8") as file:
            self.assertRegex(file.read(), 'content="text/html; charset=us-ascii"')

    def test_get_encoding(self):

        doc = DocumentHelper.create_document_fill_with_dummy_text()

        html_fixed_save_options = aw.saving.HtmlFixedSaveOptions()
        html_fixed_save_options.encoding = "utf-16"

        doc.save(ARTIFACTS_DIR + "HtmlFixedSaveOptions.get_encoding.html", html_fixed_save_options)

    def test_export_embedded_css(self):

        for export_embedded_css in (True, False):
            with self.subTest(export_embedded_css=export_embedded_css):
                #ExStart
                #ExFor:HtmlFixedSaveOptions.export_embedded_css
                #ExSummary:Shows how to determine where to store CSS stylesheets when exporting a document to Html.
                doc = aw.Document(MY_DIR + "Rendering.docx")

                # When we export a document to html, Aspose.Words will also create a CSS stylesheet to format the document with.
                # Setting the "html_fixed_save_options" flag to "True" save the CSS stylesheet to a .css file,
                # and link to the file from the html document using a <link> element.
                # Setting the flag to "False" will embed the CSS stylesheet within the Html document,
                # which will create only one file instead of two.
                html_fixed_save_options = aw.saving.HtmlFixedSaveOptions()
                html_fixed_save_options.export_embedded_css = export_embedded_css

                doc.save(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_css.html", html_fixed_save_options)

                with open(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_css.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if export_embedded_css:
                    self.assertIn('<style type="text/css">', out_doc_contents)
                    self.assertFalse(os.path.exists(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_css/styles.css"))
                else:
                    self.assertIn(
                        '<link rel="stylesheet" type="text/css" href="HtmlFixedSaveOptions.export_embedded_css/styles.css" media="all" />',
                        out_doc_contents)
                    self.assertTrue(os.path.exists(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_css/styles.css"))

                #ExEnd

                shutil.rmtree(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_css")

    def test_export_embedded_fonts(self):

        for export_embedded_fonts in (True, False):
            with self.subTest(export_embedded_fonts=export_embedded_fonts):
                #ExStart
                #ExFor:HtmlFixedSaveOptions.export_embedded_fonts
                #ExSummary:Shows how to determine where to store embedded fonts when exporting a document to Html.
                doc = aw.Document(MY_DIR + "Embedded font.docx")

                # When we export a document with embedded fonts to .html,
                # Aspose.Words can place the fonts in two possible locations.
                # Setting the "export_embedded_fonts" flag to "True" will store the raw data for embedded fonts within the CSS stylesheet,
                # in the "url" property of the "@font-face" rule. This may create a huge CSS stylesheet file
                # and reduce the number of external files that this HTML conversion will create.
                # Setting this flag to "False" will create a file for each font.
                # The CSS stylesheet will link to each font file using the "url" property of the "@font-face" rule.
                html_fixed_save_options = aw.saving.HtmlFixedSaveOptions()
                html_fixed_save_options.export_embedded_fonts = export_embedded_fonts

                doc.save(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_fonts.html", html_fixed_save_options)

                with open(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_fonts/styles.css", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if export_embedded_fonts:
                    self.assertRegex(out_doc_contents,
                        "@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], url[(].+[)] format[(]'woff'[)]; }")
                    self.assertEqual(0, len(glob.glob(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_fonts/*.woff")))
                else:
                    self.assertRegex(out_doc_contents,
                        "@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], url[(]'font001[.]woff'[)] format[(]'woff'[)]; }")
                    self.assertEqual(2, len(glob.glob(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_fonts/*.woff")))

                #ExEnd

                shutil.rmtree(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_fonts")

    def test_export_embedded_images(self):

        for export_images in (True, False):
            with self.subTest(export_images=export_images):
                #ExStart
                #ExFor:HtmlFixedSaveOptions.export_embedded_images
                #ExSummary:Shows how to determine where to store images when exporting a document to Html.
                doc = aw.Document(MY_DIR + "Images.docx")

                # When we export a document with embedded images to .html,
                # Aspose.Words can place the images in two possible locations.
                # Setting the "export_embedded_images" flag to "True" will store the raw data
                # for all images within the output HTML document, in the "src" attribute of <image> tags.
                # Setting this flag to "False" will create an image file in the local file system for every image,
                # and store all these files in a separate folder.
                html_fixed_save_options = aw.saving.HtmlFixedSaveOptions()
                html_fixed_save_options.export_embedded_images = export_images

                doc.save(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_images.html", html_fixed_save_options)

                with open(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_images.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if export_images:
                    self.assertFalse(os.path.exists(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_images/image001.jpeg"))
                    self.assertRegex(out_doc_contents,
                        '<img class="awimg" style="left:0pt; top:0pt; width:493.1pt; height:300.55pt;" src=".+" />')
                else:
                    self.assertTrue(os.path.exists(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_images/image001.jpeg"))
                    self.assertIn(
                        '<img class="awimg" style="left:0pt; top:0pt; width:493.1pt; height:300.55pt;" ' +
                        'src="HtmlFixedSaveOptions.export_embedded_images/image001.jpeg" />',
                        out_doc_contents)

                #ExEnd

                shutil.rmtree(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_images")

    def test_export_embedded_svgs(self):

        for export_svgs in (True, False):
            with self.subTest(export_svgs=export_svgs):
                #ExStart
                #ExFor:HtmlFixedSaveOptions.export_embedded_svg
                #ExSummary:Shows how to determine where to store SVG objects when exporting a document to Html.
                doc = aw.Document(MY_DIR + "Images.docx")

                # When we export a document with SVG objects to .html,
                # Aspose.Words can place these objects in two possible locations.
                # Setting the "export_embedded_svg" flag to "True" will embed all SVG object raw data
                # within the output HTML, inside <image> tags.
                # Setting this flag to "False" will create a file in the local file system for each SVG object.
                # The HTML will link to each file using the "data" attribute of an <object> tag.
                html_fixed_save_options = aw.saving.HtmlFixedSaveOptions()
                html_fixed_save_options.export_embedded_svg = export_svgs

                doc.save(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_svgs.html", html_fixed_save_options)

                with open(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_svgs.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if export_svgs:
                    self.assertFalse(os.path.exists(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_svgs/svg001.svg"))
                    self.assertRegex(out_doc_contents, '<image id="image004" xlink:href=.+/>')
                else:
                    self.assertTrue(os.path.exists(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_svgs/svg001.svg"))
                    self.assertRegex(out_doc_contents,
                        '<object type="image/svg[+]xml" data="HtmlFixedSaveOptions.export_embedded_svgs/svg001[.]svg"></object>')

                #ExEnd

                shutil.rmtree(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_embedded_svgs")

    def test_export_form_fields(self):

        for export_form_fields in (True, False):
            with self.subTest(export_form_fields=export_form_fields):
                #ExStart
                #ExFor:HtmlFixedSaveOptions.export_form_fields
                #ExSummary:Shows how to export form fields to Html.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.insert_check_box("CheckBox", False, 15)

                # When we export a document with form fields to .html,
                # there are two ways in which Aspose.Words can export form fields.
                # Setting the "export_form_fields" flag to "True" will export them as interactive objects.
                # Setting this flag to "False" will display form fields as plain text.
                # This will freeze them at their current value, and prevent the reader of our HTML document
                # from being able to interact with them.
                html_fixed_save_options = aw.saving.HtmlFixedSaveOptions()
                html_fixed_save_options.export_form_fields = export_form_fields

                doc.save(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_form_fields.html", html_fixed_save_options)

                with open(ARTIFACTS_DIR + "HtmlFixedSaveOptions.export_form_fields.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if export_form_fields:
                    self.assertRegex(out_doc_contents,
                        '<a name="CheckBox" style="left:0pt; top:0pt;"></a>' +
                        '<input style="position:absolute; left:0pt; top:0pt;" type="checkbox" name="CheckBox" />')
                else:
                    self.assertRegex(out_doc_contents,
                        '<a name="CheckBox" style="left:0pt; top:0pt;"></a>' +
                        '<div class="awdiv" style="left:0.8pt; top:0.8pt; width:14.25pt; height:14.25pt; border:solid 0.75pt #000000;"')

                #ExEnd

    def test_add_css_class_names_prefix(self):

        #ExStart
        #ExFor:HtmlFixedSaveOptions.css_class_names_prefix
        #ExFor:HtmlFixedSaveOptions.save_font_face_css_separately
        #ExSummary:Shows how to place CSS into a separate file and add a prefix to all of its CSS class names.
        doc = aw.Document(MY_DIR + "Bookmarks.docx")

        html_fixed_save_options = aw.saving.HtmlFixedSaveOptions()
        html_fixed_save_options.css_class_names_prefix = "myprefix"
        html_fixed_save_options.save_font_face_css_separately = True

        doc.save(ARTIFACTS_DIR + "HtmlFixedSaveOptions.add_css_class_names_prefix.html", html_fixed_save_options)

        with open(ARTIFACTS_DIR + "HtmlFixedSaveOptions.add_css_class_names_prefix.html", "rt", encoding="utf-8") as file:
            out_doc_contents = file.read()

        self.assertRegex(out_doc_contents,
            '<div class="myprefixdiv myprefixpage" style="width:595[.]3pt; height:841[.]9pt;">' +
            '<div class="myprefixdiv" style="left:85[.]05pt; top:36pt; clip:rect[(]0pt,510[.]25pt,74[.]95pt,-85.05pt[)];">' +
            '<span class="myprefixspan myprefixtext001" style="font-size:11pt; left:294[.]73pt; top:0[.]36pt;">')

        with open(ARTIFACTS_DIR + "HtmlFixedSaveOptions.add_css_class_names_prefix/styles.css", "rt", encoding="utf-8") as file:
            out_doc_contents = file.read()

        self.assertRegex(out_doc_contents,
            ".myprefixdiv { position:absolute; } " +
            ".myprefixspan { position:absolute; white-space:pre; color:#000000; font-size:12pt; }")
        #ExEnd

    def test_horizontal_alignment(self):

        for page_horizontal_alignment in (aw.saving.HtmlFixedPageHorizontalAlignment.CENTER,
                                          aw.saving.HtmlFixedPageHorizontalAlignment.LEFT,
                                          aw.saving.HtmlFixedPageHorizontalAlignment.RIGHT):
            with self.subTest(page_horizontal_alignment=page_horizontal_alignment):
                #ExStart
                #ExFor:HtmlFixedSaveOptions.page_horizontal_alignment
                #ExFor:HtmlFixedPageHorizontalAlignment
                #ExSummary:Shows how to set the horizontal alignment of pages when saving a document to HTML.
                doc = aw.Document(MY_DIR + "Rendering.docx")

                html_fixed_save_options = aw.saving.HtmlFixedSaveOptions()
                html_fixed_save_options.page_horizontal_alignment = page_horizontal_alignment

                doc.save(ARTIFACTS_DIR + "HtmlFixedSaveOptions.horizontal_alignment.html", html_fixed_save_options)

                with open(ARTIFACTS_DIR + "HtmlFixedSaveOptions.horizontal_alignment/styles.css", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if page_horizontal_alignment == aw.saving.HtmlFixedPageHorizontalAlignment.CENTER:
                    self.assertRegex(out_doc_contents,
                        "[.]awpage { position:relative; border:solid 1pt black; margin:10pt auto 10pt auto; overflow:hidden; }")

                elif page_horizontal_alignment == aw.saving.HtmlFixedPageHorizontalAlignment.LEFT:
                    self.assertRegex(out_doc_contents,
                        "[.]awpage { position:relative; border:solid 1pt black; margin:10pt auto 10pt 10pt; overflow:hidden; }")

                elif page_horizontal_alignment == aw.saving.HtmlFixedPageHorizontalAlignment.RIGHT:
                    self.assertRegex(out_doc_contents,
                        "[.]awpage { position:relative; border:solid 1pt black; margin:10pt 10pt 10pt auto; overflow:hidden; }")

                #ExEnd

    def test_page_margins(self):

        #ExStart
        #ExFor:HtmlFixedSaveOptions.page_margins
        #ExSummary:Shows how to adjust page margins when saving a document to HTML.
        doc = aw.Document(MY_DIR + "Document.docx")

        save_options = aw.saving.HtmlFixedSaveOptions()
        save_options.page_margins = 15

        doc.save(ARTIFACTS_DIR + "HtmlFixedSaveOptions.page_margins.html", save_options)

        with open(ARTIFACTS_DIR + "HtmlFixedSaveOptions.page_margins/styles.css", "rt", encoding="utf-8") as file:
            out_doc_contents = file.read()

        self.assertRegex(out_doc_contents,
            "[.]awpage { position:relative; border:solid 1pt black; margin:15pt auto 15pt auto; overflow:hidden; }")
        #ExEnd

    def test_page_margins_exception(self):

        save_options = aw.saving.HtmlFixedSaveOptions()
        with self.assertRaises(Exception):
            save_options.page_margins = -1

    def test_optimize_graphics_output(self):

        for optimize_output in (False, True):
            with self.subTest(optimize_output=optimize_output):
                #ExStart
                #ExFor:HtmlFixedSaveOptions.optimize_output
                #ExSummary:Shows how to simplify a document when saving it to HTML by removing various redundant objects.
                doc = aw.Document(MY_DIR + "Rendering.docx")

                save_options = aw.saving.HtmlFixedSaveOptions()
                save_options.optimize_output = optimize_output

                doc.save(ARTIFACTS_DIR + "HtmlFixedSaveOptions.optimize_graphics_output.html", save_options)

                # The size of the optimized version of the document is almost a third of the size of the unoptimized document.
                if optimize_output:
                    self.assertAlmostEqual(57220,
                        os.path.getsize(ARTIFACTS_DIR + "HtmlFixedSaveOptions.optimize_graphics_output.html"), delta=200)
                else:
                    self.assertAlmostEqual(160535,
                        os.path.getsize(ARTIFACTS_DIR + "HtmlFixedSaveOptions.optimize_graphics_output.html"), delta=200)
                #ExEnd

    def test_using_machine_fonts(self):

        for use_target_machine_fonts in (False, True):
            with self.subTest(use_target_machine_fonts=use_target_machine_fonts):
                #ExStart
                #ExFor:ExportFontFormat
                #ExFor:HtmlFixedSaveOptions.font_format
                #ExFor:HtmlFixedSaveOptions.use_target_machine_fonts
                #ExSummary:Shows how use fonts only from the target machine when saving a document to HTML.
                doc = aw.Document(MY_DIR + "Bullet points with alternative font.docx")

                save_options = aw.saving.HtmlFixedSaveOptions()
                save_options.export_embedded_css = True
                save_options.use_target_machine_fonts = use_target_machine_fonts
                save_options.font_format = aw.saving.ExportFontFormat.TTF
                save_options.export_embedded_fonts = False

                doc.save(ARTIFACTS_DIR + "HtmlFixedSaveOptions.using_machine_fonts.html", save_options)

                with open(ARTIFACTS_DIR + "HtmlFixedSaveOptions.using_machine_fonts.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if use_target_machine_fonts:
                    self.assertNotRegex(out_doc_contents, "@font-face")
                else:
                    self.assertIn(
                        "@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local('☺'), " +
                        "url('HtmlFixedSaveOptions.using_machine_fonts/font001.ttf') format('truetype'); }",
                        out_doc_contents)
                #ExEnd

    ##ExStart
    ##ExFor:IResourceSavingCallback
    ##ExFor:IResourceSavingCallback.resource_saving(ResourceSavingArgs)
    ##ExFor:ResourceSavingArgs
    ##ExFor:ResourceSavingArgs.document
    ##ExFor:ResourceSavingArgs.resource_file_name
    ##ExFor:ResourceSavingArgs.resource_file_uri
    ##ExSummary:Shows how to use a callback to track external resources created while converting a document to HTML.
    #def test_resource_saving_callback(self):

    #    doc = aw.Document(MY_DIR + "Bullet points with alternative font.docx")

    #    callback = ExHtmlFixedSaveOptions.FontSavingCallback()

    #    save_options = aw.saving.HtmlFixedSaveOptions
    #    save_options.resource_saving_callback = callback

    #    doc.save(ARTIFACTS_DIR + "HtmlFixedSaveOptions.using_machine_fonts.html", save_options)

    #    print(callback.text)
    #    self._test_resource_saving_callback(callback) #ExSkip

    #class FontSavingCallback(aw.saving.IResourceSavingCallback):

    #    def __init__(self):
    #        self.text = ""

    #    def resource_saving(self, args: aw.saving.ResourceSavingArgs):
    #        """Called when Aspose.Words saves an external resource to fixed page HTML or SVG."""

    #        self.text += f"Original document URI:\t{args.document.original_file_name}\n"
    #        self.text += f"Resource being saved:\t{args.resource_file_name}\n"
    #        self.text += f"Full uri after saving:\t{args.resource_file_uri}\n\n"

    ##ExEnd

    #def _test_resource_saving_callback(self, callback):

    #    self.assertIn("font001.woff", callback.text)
    #    self.assertIn("styles.css", callback.text)

    ##ExStart
    ##ExFor:HtmlFixedSaveOptions
    ##ExFor:HtmlFixedSaveOptions.resource_saving_callback
    ##ExFor:HtmlFixedSaveOptions.resources_folder
    ##ExFor:HtmlFixedSaveOptions.resources_folder_alias
    ##ExFor:HtmlFixedSaveOptions.save_format
    ##ExFor:HtmlFixedSaveOptions.show_page_border
    ##ExFor:IResourceSavingCallback
    ##ExFor:IResourceSavingCallback.resource_saving(ResourceSavingArgs)
    ##ExFor:ResourceSavingArgs.keep_resource_stream_open
    ##ExFor:ResourceSavingArgs.resource_stream
    ##ExSummary:Shows how to use a callback to print the URIs of external resources created while converting a document to HTML.
    #def test_html_fixed_resource_folder(self):

    #    doc = aw.Document(MY_DIR + "Rendering.docx")

    #    callback = ExHtmlFixedSaveOptions.ResourceUriPrinter()

    #    options = aw.saving.HtmlFixedSaveOptions()
    #    options.save_format = aw.SaveFormat.HTML_FIXED
    #    options.export_embedded_images = False
    #    options.resources_folder = ARTIFACTS_DIR + "HtmlFixedResourceFolder"
    #    options.resources_folder_alias = ARTIFACTS_DIR + "HtmlFixedResourceFolderAlias"
    #    options.show_page_border = False
    #    options.resource_saving_callback = callback

    #    # A folder specified by resources_folder_alias will contain the resources instead of ResourcesFolder.
    #    # We must ensure the folder exists before the streams can put their resources into it.
    #    os.mkdir(options.resources_folder_alias)

    #    doc.save(ARTIFACTS_DIR + "HtmlFixedSaveOptions.html_fixed_resource_folder.html", options)

    #    print(callback.text)

    #    resource_files = glog.glob(ARTIFACTS_DIR + "HtmlFixedResourceFolderAlias/*.*")

    #    self.assertFalse(os.path.exists(ARTIFACTS_DIR + "HtmlFixedResourceFolder"))
    #    self.assertEqual(6, len([f for f in resource_files if f.endswith(".jpeg") or f.endswith(".png") or f.endswith(".css")]))
    #    self._test_html_fixed_resource_folder(callback) #ExSkip

    #class ResourceUriPrinter(aw.saving.IResourceSavingCallback):
    #    """Counts and prints URIs of resources contained by as they are converted to fixed HTML."""

    #    def __init__(self):
    #        self.saved_resource_count = 0
    #        self.text = ""

    #    def resource_saving(self, args: aw.saving.ResourceSavingArgs) -> aw.saving.IResourceSavingCallback:

    #        self.saved_resource_count += 1

    #        # If we set a folder alias in the SaveOptions object, we will be able to print it from here.
    #        self.text += f'Resource #{self.saved_resource_count} "{args.resource_file_name}"\n'

    #        extension = os.path.splitext(args.resource_file_name)[-1]
    #        if extension in ("ttf", "woff"):
    #            # By default, 'resource_file_uri' uses system folder for fonts.
    #            # To avoid problems in other platforms you must explicitly specify the path for the fonts.
    #            args.resource_file_uri = ARTIFACTS_DIR + os.path.sep + args.resource_file_name

    #        self.text += "\t" + args.resource_file_uri + "\n"

    #        # If we have specified a folder in the "resources_folder_alias" property,
    #        # we will also need to redirect each stream to put its resource in that folder.
    #        args.resource_stream = open(args.resource_file_uri, "wb")
    #        args.keep_resource_stream_open = False

    ##ExEnd

    #def _test_html_fixed_resource_folder(self, callback):

    #    self.assertEqual(16, callback.text.count("Resource #"))
