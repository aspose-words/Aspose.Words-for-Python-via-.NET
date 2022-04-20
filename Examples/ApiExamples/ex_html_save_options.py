# Copyright (c) 2001-2022 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import io
import os
import glob
import textwrap
import shutil
import unittest

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, IMAGE_DIR
from document_helper import DocumentHelper

class ExHtmlSaveOptions(ApiExampleBase):

    def test_export_page_margins_epub(self):

        for save_format in (aw.SaveFormat.HTML,
                            aw.SaveFormat.MHTML,
                            aw.SaveFormat.EPUB):
            with self.subTest(save_format=save_format):
                doc = aw.Document(MY_DIR + "TextBoxes.docx")

                save_options = aw.saving.HtmlSaveOptions()
                save_options.save_format = save_format
                save_options.export_page_margins = True

                doc.save(
                    ARTIFACTS_DIR + "HtmlSaveOptions.export_page_margins_epub" +
                    aw.FileFormatUtil.save_format_to_extension(save_format), save_options)

    def test_export_office_math_epub(self):

        parameters = [
            (aw.SaveFormat.HTML, aw.saving.HtmlOfficeMathOutputMode.IMAGE),
            (aw.SaveFormat.MHTML, aw.saving.HtmlOfficeMathOutputMode.MATH_ML),
            (aw.SaveFormat.EPUB, aw.saving.HtmlOfficeMathOutputMode.TEXT)]

        for save_format, output_mode in parameters:
            with self.subTest(save_format=save_format, output_mode=output_mode):
                doc = aw.Document(MY_DIR + "Office math.docx")

                save_options = aw.saving.HtmlSaveOptions()
                save_options.office_math_output_mode = output_mode

                doc.save(
                    ARTIFACTS_DIR + "HtmlSaveOptions.export_office_math_epub" +
                    aw.FileFormatUtil.save_format_to_extension(save_format), save_options)

    def test_export_text_box_as_svg_epub(self):

        parameters = [
            (aw.SaveFormat.HTML, True, "TextBox as svg (html)"),
            (aw.SaveFormat.EPUB, True, "TextBox as svg (epub)"),
            (aw.SaveFormat.MHTML, False, "TextBox as img (mhtml)")]

        for save_format, is_text_box_as_svg, description in parameters:
            with self.subTest(description=description):
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                textbox = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 300, 100)
                builder.move_to(textbox.first_paragraph)
                builder.write("Hello world!")

                save_options = aw.saving.HtmlSaveOptions(save_format)
                save_options.export_text_box_as_svg = is_text_box_as_svg

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.export_text_box_as_svg_epub" + aw.FileFormatUtil.save_format_to_extension(save_format), save_options)

                if save_format == aw.SaveFormat.HTML:
                    dir_files = glob.glob(ARTIFACTS_DIR + "**/HtmlSaveOptions.export_text_box_as_svg_epub.001.png", recursive=True)
                    self.assertEqual(0, len(dir_files))

                elif save_format == aw.SaveFormat.EPUB:
                    dir_files = glob.glob(ARTIFACTS_DIR + "**/HtmlSaveOptions.export_text_box_as_svg_epub.001.png", recursive=True)
                    self.assertEqual(0, len(dir_files))

                elif save_format == aw.SaveFormat.MHTML:
                    dir_files = glob.glob(ARTIFACTS_DIR + "**/HtmlSaveOptions.export_text_box_as_svg_epub.001.png", recursive=True)
                    self.assertEqual(0, len(dir_files))

    def test_control_list_labels_export(self):

        for how_export_list_labels in (aw.saving.ExportListLabels.AUTO,
                                       aw.saving.ExportListLabels.AS_INLINE_TEXT,
                                       aw.saving.ExportListLabels.BY_HTML_TAGS):
            with self.subTest(how_export_list_labels=how_export_list_labels):
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                bulleted_list = doc.lists.add(aw.lists.ListTemplate.BULLET_DEFAULT)
                builder.list_format.list = bulleted_list
                builder.paragraph_format.left_indent = 72
                builder.writeln("Bulleted list item 1.")
                builder.writeln("Bulleted list item 2.")
                builder.paragraph_format.clear_formatting()

                save_options = aw.saving.HtmlSaveOptions(aw.SaveFormat.HTML)

                # 'ExportListLabels.AUTO' - this option uses <ul> and <ol> tags are used for list label representation if it does not cause formatting loss,
                # otherwise HTML <p> tag is used. This is also the default value.
                # 'ExportListLabels.AS_INLINE_TEXT' - using this option the <p> tag is used for any list label representation.
                # 'ExportListLabels.BY_HTML_TAGS' - The <ul> and <ol> tags are used for list label representation. Some formatting loss is possible.
                save_options.export_list_labels = how_export_list_labels

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.control_list_labels_export.html", save_options)

    def test_export_url_for_linked_image(self):

        for export in (True, False):
            with self.subTest(export=export):
                doc = aw.Document(MY_DIR + "Linked image.docx")

                save_options = aw.saving.HtmlSaveOptions()
                save_options.export_original_url_for_linked_images = export

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.export_url_for_linked_image.html", save_options)

                dir_files = glob.glob(ARTIFACTS_DIR + "**/HtmlSaveOptions.export_url_for_linked_image.001.png", recursive=True)

                DocumentHelper.find_text_in_file(ARTIFACTS_DIR + "HtmlSaveOptions.export_url_for_linked_image.html",
                    "<img src=\"http://www.aspose.com/images/aspose-logo.gif\"" if not dir_files else "<img src=\"HtmlSaveOptions.export_url_for_linked_image.001.png\"")

    def test_export_roundtrip_information(self):

        doc = aw.Document(MY_DIR + "TextBoxes.docx")
        save_options = aw.saving.HtmlSaveOptions()
        save_options.export_roundtrip_information = True

        doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.export_roundtrip_information.html", save_options)

    def test_roundtrip_information_defaul_value(self):

        save_options = aw.saving.HtmlSaveOptions(aw.SaveFormat.HTML)
        self.assertEqual(True, save_options.export_roundtrip_information)

        save_options = aw.saving.HtmlSaveOptions(aw.SaveFormat.MHTML)
        self.assertEqual(False, save_options.export_roundtrip_information)

        save_options = aw.saving.HtmlSaveOptions(aw.SaveFormat.EPUB)
        self.assertEqual(False, save_options.export_roundtrip_information)

    def test_external_resource_saving_config(self):

        for filepath in glob.glob(ARTIFACTS_DIR + "Resources/HtmlSaveOptions.external_resource_saving_config*"):
            os.remove(filepath)

        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.HtmlSaveOptions()
        save_options.css_style_sheet_type = aw.saving.CssStyleSheetType.EXTERNAL
        save_options.export_font_resources = True
        save_options.resource_folder = "Resources"
        save_options.resource_folder_alias = "https://www.aspose.com/"

        doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.external_resource_saving_config.html", save_options)

        image_files = glob.glob(ARTIFACTS_DIR + "Resources/HtmlSaveOptions.external_resource_saving_config*.png")
        self.assertEqual(8, len(image_files))

        font_files = glob.glob(ARTIFACTS_DIR + "Resources/HtmlSaveOptions.external_resource_saving_config*.ttf")
        self.assertEqual(10, len(font_files))

        css_files = glob.glob(ARTIFACTS_DIR + "Resources/HtmlSaveOptions.external_resource_saving_config*.css")
        self.assertEqual(1, len(css_files))

        DocumentHelper.find_text_in_file(ARTIFACTS_DIR + "HtmlSaveOptions.external_resource_saving_config.html",
            "<link href=\"https://www.aspose.com/HtmlSaveOptions.external_resource_saving_config.css\"")

    def test_convert_fonts_as_base64(self):

        doc = aw.Document(MY_DIR + "TextBoxes.docx")

        save_options = aw.saving.HtmlSaveOptions()
        save_options.css_style_sheet_type = aw.saving.CssStyleSheetType.EXTERNAL
        save_options.resource_folder = "Resources"
        save_options.export_font_resources = True
        save_options.export_fonts_as_base64 = True

        doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.convert_fonts_as_base64.html", save_options)

    def test_html5_support(self):

        for html_version in (aw.saving.HtmlVersion.HTML5,
                             aw.saving.HtmlVersion.XHTML):
            with self.subTest(html_version=html_version):
                doc = aw.Document(MY_DIR + "Document.docx")

                save_options = aw.saving.HtmlSaveOptions()
                save_options.html_version = html_version

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.html5_support.html", save_options)

    def test_export_fonts(self):

        for export_as_base64 in (False, True):
            with self.subTest(export_as_base64=export_as_base64):
                fonts_folder = ARTIFACTS_DIR + "HtmlSaveOptions.export_fonts.resources/"

                doc = aw.Document(MY_DIR + "Document.docx")

                save_options = aw.saving.HtmlSaveOptions()
                save_options.export_font_resources = True
                save_options.fonts_folder = fonts_folder
                save_options.export_fonts_as_base64 = export_as_base64

                if export_as_base64:
                    doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.export_fonts.True.html", save_options)
                    self.assertFalse(os.path.exists(fonts_folder))
                else:
                    doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.export_fonts.False.html", save_options)

                    self.assertTrue(os.path.exists(fonts_folder + "HtmlSaveOptions.export_fonts.False.times.ttf"))
                    shutil.rmtree(fonts_folder)

    def test_resource_folder_priority(self):

        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.HtmlSaveOptions()
        save_options.css_style_sheet_type = aw.saving.CssStyleSheetType.EXTERNAL
        save_options.export_font_resources = True
        save_options.resource_folder = ARTIFACTS_DIR + "Resources"
        save_options.resource_folder_alias = "http://example.com/resources"

        doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.resource_folder_priority.html", save_options)

        self.assertTrue(os.path.exists(ARTIFACTS_DIR + "Resources/HtmlSaveOptions.resource_folder_priority.001.png"))
        self.assertTrue(os.path.exists(ARTIFACTS_DIR + "Resources/HtmlSaveOptions.resource_folder_priority.002.png"))
        self.assertTrue(os.path.exists(ARTIFACTS_DIR + "Resources/HtmlSaveOptions.resource_folder_priority.arial.ttf"))
        self.assertTrue(os.path.exists(ARTIFACTS_DIR + "Resources/HtmlSaveOptions.resource_folder_priority.css"))

    def test_resource_folder_low_priority(self):

        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.HtmlSaveOptions()
        save_options.css_style_sheet_type = aw.saving.CssStyleSheetType.EXTERNAL
        save_options.export_font_resources = True
        save_options.fonts_folder = ARTIFACTS_DIR + "Fonts"
        save_options.images_folder = ARTIFACTS_DIR + "Images"
        save_options.resource_folder = ARTIFACTS_DIR + "Resources"
        save_options.resource_folder_alias = "http://example.com/resources"

        doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.resource_folder_low_priority.html", save_options)

        self.assertTrue(os.path.exists(ARTIFACTS_DIR + "Images/HtmlSaveOptions.resource_folder_low_priority.001.png"))
        self.assertTrue(os.path.exists(ARTIFACTS_DIR + "Images/HtmlSaveOptions.resource_folder_low_priority.002.png"))
        self.assertTrue(os.path.exists(ARTIFACTS_DIR + "Fonts/HtmlSaveOptions.resource_folder_low_priority.arial.ttf"))
        self.assertTrue(os.path.exists(ARTIFACTS_DIR + "Resources/HtmlSaveOptions.resource_folder_low_priority.css"))

    def test_svg_metafile_format(self):

        builder = aw.DocumentBuilder()

        builder.write("Here is an SVG image: ")
        builder.insert_html("""<svg height='210' width='500'>
            <polygon points='100,10 40,198 190,78 10,78 160,198'
                style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />
            </svg>""")

        save_options = aw.saving.HtmlSaveOptions()
        save_options.metafile_format = aw.saving.HtmlMetafileFormat.PNG

        builder.document.save(ARTIFACTS_DIR + "HtmlSaveOptions.svg_metafile_format.html", save_options)

    def test_png_metafile_format(self):

        builder = aw.DocumentBuilder()

        builder.write("Here is an Png image: ")
        builder.insert_html("""<svg height='210' width='500'>
            <polygon points='100,10 40,198 190,78 10,78 160,198'
                style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />
            </svg>""")

        save_options = aw.saving.HtmlSaveOptions()
        save_options.metafile_format = aw.saving.HtmlMetafileFormat.PNG

        builder.document.save(ARTIFACTS_DIR + "HtmlSaveOptions.png_metafile_format.html", save_options)

    def test_emf_or_wmf_metafile_format(self):

        builder = aw.DocumentBuilder()

        builder.write("Here is an image as is: ")
        builder.insert_html(textwrap.dedent("""<img src=""data:image/png;base64,
            iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAABGdBTUEAALGP
            C/xhBQAAAAlwSFlzAAALEwAACxMBAJqcGAAAAAd0SU1FB9YGARc5KB0XV+IA
            AAAddEVYdENvbW1lbnQAQ3JlYXRlZCB3aXRoIFRoZSBHSU1Q72QlbgAAAF1J
            REFUGNO9zL0NglAAxPEfdLTs4BZM4DIO4C7OwQg2JoQ9LE1exdlYvBBeZ7jq
            ch9//q1uH4TLzw4d6+ErXMMcXuHWxId3KOETnnXXV6MJpcq2MLaI97CER3N0
            vr4MkhoXe0rZigAAAABJRU5ErkJggg=="" alt=""Red dot"" />"""))

        save_options = aw.saving.HtmlSaveOptions()
        save_options.metafile_format = aw.saving.HtmlMetafileFormat.EMF_OR_WMF

        builder.document.save(ARTIFACTS_DIR + "HtmlSaveOptions.emf_or_wmf_metafile_format.html", save_options)

    def test_css_class_names_prefix(self):

        #ExStart
        #ExFor:HtmlSaveOptions.css_class_name_prefix
        #ExSummary:Shows how to save a document to HTML, and add a prefix to all of its CSS class names.
        doc = aw.Document(MY_DIR + "Paragraphs.docx")

        save_options = aw.saving.HtmlSaveOptions()
        save_options.css_style_sheet_type = aw.saving.CssStyleSheetType.EXTERNAL
        save_options.css_class_name_prefix = "myprefix-"

        doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.css_class_name_prefix.html", save_options)

        with open(ARTIFACTS_DIR + "HtmlSaveOptions.css_class_name_prefix.html", "rt", encoding="utf-8") as file:
            out_doc_contents = file.read()

        self.assertIn("<p class=\"myprefix-Header\">", out_doc_contents)
        self.assertIn("<p class=\"myprefix-Footer\">", out_doc_contents)

        with open(ARTIFACTS_DIR + "HtmlSaveOptions.css_class_name_prefix.css", "rt", encoding="utf-8") as file:
            out_doc_contents = file.read()

        self.assertIn(
            ".myprefix-Footer { margin-bottom:0pt; line-height:normal; font-family:Arial; font-size:11pt }\n" +
            ".myprefix-Header { margin-bottom:0pt; line-height:normal; font-family:Arial; font-size:11pt }\n",
            out_doc_contents)
        #ExEnd

    def test_css_class_names_not_valid_prefix(self):

        save_options = aw.saving.HtmlSaveOptions()
        with self.assertRaises(Exception, msg="The class name prefix must be a valid CSS identifier."):
            save_options.css_class_name_prefix = "@%-"

    def test_css_class_names_null_prefix(self):

        doc = aw.Document(MY_DIR + "Paragraphs.docx")

        save_options = aw.saving.HtmlSaveOptions()
        save_options.css_style_sheet_type = aw.saving.CssStyleSheetType.EMBEDDED
        save_options.css_class_name_prefix = None

        doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.css_class_names_null_prefix.html", save_options)

    def test_content_id_scheme(self):

        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.HtmlSaveOptions(aw.SaveFormat.MHTML)
        save_options.pretty_format = True
        save_options.export_cid_urls_for_mhtml_resources = True

        doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.content_id_scheme.mhtml", save_options)

    @unittest.skip("Bug")
    def test_resolve_font_names(self):

        for resolve_font_names in (False, True):
            with self.subTest(resolve_font_names=resolve_font_names):
                #ExStart
                #ExFor:HtmlSaveOptions.resolve_font_names
                #ExSummary:Shows how to resolve all font names before writing them to HTML.
                doc = aw.Document(MY_DIR + "Missing font.docx")

                # This document contains text that names a font that we do not have.
                self.assertIsNotNone(doc.font_infos.get_by_name("28 Days Later"))

                # If we have no way of getting this font, and we want to be able to display all the text
                # in this document in an output HTML, we can substitute it with another font.
                font_settings = aw.fonts.FontSettings()
                font_settings.substitution_settings.default_font_substitution.default_font_name = "Arial"
                font_settings.substitution_settings.default_font_substitution.enabled = True

                doc.font_settings = font_settings

                save_options = aw.saving.HtmlSaveOptions(aw.SaveFormat.HTML)

                # By default, this option is set to 'False' and Aspose.Words writes font names as specified in the source document
                save_options.resolve_font_names = resolve_font_names

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.resolve_font_names.html", save_options)

                with open(ARTIFACTS_DIR + "HtmlSaveOptions.resolve_font_names.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if resolve_font_names:
                    self.assertIn("<span style=\"font-family:Arial\">", out_doc_contents)
                else:
                    self.assertIn("<span style=\"font-family:'28 Days Later'\">", out_doc_contents)
                #ExEnd

    def test_heading_levels(self):

        #ExStart
        #ExFor:HtmlSaveOptions.document_split_heading_level
        #ExSummary:Shows how to split an output HTML document by headings into several parts.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Every paragraph that we format using a "Heading" style can serve as a heading.
        # Each heading may also have a heading level, determined by the number of its heading style.
        # The headings below are of levels 1-3.
        builder.paragraph_format.style = builder.document.styles.get_by_name("Heading 1")
        builder.writeln("Heading #1")
        builder.paragraph_format.style = builder.document.styles.get_by_name("Heading 2")
        builder.writeln("Heading #2")
        builder.paragraph_format.style = builder.document.styles.get_by_name("Heading 3")
        builder.writeln("Heading #3")
        builder.paragraph_format.style = builder.document.styles.get_by_name("Heading 1")
        builder.writeln("Heading #4")
        builder.paragraph_format.style = builder.document.styles.get_by_name("Heading 2")
        builder.writeln("Heading #5")
        builder.paragraph_format.style = builder.document.styles.get_by_name("Heading 3")
        builder.writeln("Heading #6")

        # Create a HtmlSaveOptions object and set the split criteria to "HEADING_PARAGRAPH".
        # These criteria will split the document at paragraphs with "Heading" styles into several smaller documents,
        # and save each document in a separate HTML file in the local file system.
        # We will also set the maximum heading level, which splits the document to 2.
        # Saving the document will split it at headings of levels 1 and 2, but not at 3 to 9.
        options = aw.saving.HtmlSaveOptions()
        options.document_split_criteria = aw.saving.DocumentSplitCriteria.HEADING_PARAGRAPH
        options.document_split_heading_level = 2

        # Our document has four headings of levels 1 - 2. One of those headings will not be
        # a split point since it is at the beginning of the document.
        # The saving operation will split our document at three places, into four smaller documents.
        doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.heading_levels.html", options)

        doc = aw.Document(ARTIFACTS_DIR + "HtmlSaveOptions.heading_levels.html")

        self.assertEqual("Heading #1", doc.get_text().strip())

        doc = aw.Document(ARTIFACTS_DIR + "HtmlSaveOptions.heading_levels-01.html")

        self.assertEqual("Heading #2\r" +
                         "Heading #3", doc.get_text().strip())

        doc = aw.Document(ARTIFACTS_DIR + "HtmlSaveOptions.heading_levels-02.html")

        self.assertEqual("Heading #4", doc.get_text().strip())

        doc = aw.Document(ARTIFACTS_DIR + "HtmlSaveOptions.heading_levels-03.html")

        self.assertEqual("Heading #5\r" +
                         "Heading #6", doc.get_text().strip())
        #ExEnd

    def test_negative_indent(self):

        for allow_negative_indent in (False, True):
            with self.subTest(allow_negative_indent=allow_negative_indent):
                #ExStart
                #ExFor:HtmlElementSizeOutputMode
                #ExFor:HtmlSaveOptions.allow_negative_indent
                #ExFor:HtmlSaveOptions.table_width_output_mode
                #ExSummary:Shows how to preserve negative indents in the output .html.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Insert a table with a negative indent, which will push it to the left past the left page boundary.
                table = builder.start_table()
                builder.insert_cell()
                builder.write("Row 1, Cell 1")
                builder.insert_cell()
                builder.write("Row 1, Cell 2")
                builder.end_table()
                table.left_indent = -36
                table.preferred_width = aw.tables.PreferredWidth.from_points(144)

                builder.insert_break(aw.BreakType.PARAGRAPH_BREAK)

                # Insert a table with a positive indent, which will push the table to the right.
                table = builder.start_table()
                builder.insert_cell()
                builder.write("Row 1, Cell 1")
                builder.insert_cell()
                builder.write("Row 1, Cell 2")
                builder.end_table()
                table.left_indent = 36
                table.preferred_width = aw.tables.PreferredWidth.from_points(144)

                # When we save a document to HTML, Aspose.Words will only preserve negative indents
                # such as the one we have applied to the first table if we set the "allow_negative_indent" flag
                # in a SaveOptions object that we will pass to "True".
                options = aw.saving.HtmlSaveOptions(aw.SaveFormat.HTML)
                options.allow_negative_indent = allow_negative_indent
                options.table_width_output_mode = aw.saving.HtmlElementSizeOutputMode.RELATIVE_ONLY

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.negative_indent.html", options)

                with open(ARTIFACTS_DIR + "HtmlSaveOptions.negative_indent.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if allow_negative_indent:
                    self.assertIn(
                        "<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:-41.65pt; border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">",
                        out_doc_contents)
                    self.assertIn(
                        "<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:30.35pt; border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">",
                        out_doc_contents)
                else:
                    self.assertIn(
                        "<table cellspacing=\"0\" cellpadding=\"0\" style=\"border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">",
                        out_doc_contents)
                    self.assertIn(
                        "<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:30.35pt; border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">",
                        out_doc_contents)

                #ExEnd

    def test_folder_alias(self):

        #ExStart
        #ExFor:HtmlSaveOptions.export_original_url_for_linked_images
        #ExFor:HtmlSaveOptions.fonts_folder
        #ExFor:HtmlSaveOptions.fonts_folder_alias
        #ExFor:HtmlSaveOptions.image_resolution
        #ExFor:HtmlSaveOptions.images_folder_alias
        #ExFor:HtmlSaveOptions.resource_folder
        #ExFor:HtmlSaveOptions.resource_folder_alias
        #ExSummary:Shows how to set folders and folder aliases for externally saved resources that Aspose.Words will create when saving a document to HTML.
        doc = aw.Document(MY_DIR + "Rendering.docx")

        options = aw.saving.HtmlSaveOptions()
        options.css_style_sheet_type = aw.saving.CssStyleSheetType.EXTERNAL
        options.export_font_resources = True
        options.image_resolution = 72
        options.font_resources_subsetting_size_threshold = 0
        options.fonts_folder = ARTIFACTS_DIR + "Fonts"
        options.images_folder = ARTIFACTS_DIR + "Images"
        options.resource_folder = ARTIFACTS_DIR + "Resources"
        options.fonts_folder_alias = "http://example.com/fonts"
        options.images_folder_alias = "http://example.com/images"
        options.resource_folder_alias = "http://example.com/resources"
        options.export_original_url_for_linked_images = True

        doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.folder_alias.html", options)
        #ExEnd

    ##ExStart
    ##ExFor:HtmlSaveOptions.export_font_resources
    ##ExFor:HtmlSaveOptions.font_saving_callback
    ##ExFor:IFontSavingCallback
    ##ExFor:IFontSavingCallback.font_saving
    ##ExFor:FontSavingArgs
    ##ExFor:FontSavingArgs.bold
    ##ExFor:FontSavingArgs.document
    ##ExFor:FontSavingArgs.font_family_name
    ##ExFor:FontSavingArgs.font_file_name
    ##ExFor:FontSavingArgs.font_stream
    ##ExFor:FontSavingArgs.is_export_needed
    ##ExFor:FontSavingArgs.is_subsetting_needed
    ##ExFor:FontSavingArgs.italic
    ##ExFor:FontSavingArgs.keep_font_stream_open
    ##ExFor:FontSavingArgs.original_file_name
    ##ExFor:FontSavingArgs.original_file_size
    ##ExSummary:Shows how to define custom logic for exporting fonts when saving to HTML.
    #def test_save_exported_fonts(self):

    #    doc = aw.Document(MY_DIR + "Rendering.docx")

    #    # Configure a SaveOptions object to export fonts to separate files.
    #    # Set a callback that will handle font saving in a custom manner.
    #    options = aw.saving.HtmlSaveOptions()
    #    options.export_font_resources = True
    #    options.font_saving_callback = ExHtmlSaveOptions.HandleFontSaving()

    #    # The callback will export .ttf files and save them alongside the output document.
    #    doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.save_exported_fonts.html", options)

    #    for font_filename in glob.glob(ARTIFACTS_DIR + "*.ttf"):
    #        print(font_filename)

    #    self.assertEqual(10, len(glob.glob(ARTIFACTS_DIR + "*.ttf")) #ExSkip

    #class HandleFontSaving(aw.saving.IFontSavingCallback):
    #    """Prints information about exported fonts and saves them in the same local system folder as their output .html."""

    #    def font_saving(self, args: aw.saving.FontSavingArgs):

    #        print(f"Font:\t{args.font_family_name}", end="")
    #        if args.bold:
    #            print(", bold", end="")
    #        if args.italic:
    #            print(", italic", end="")
    #        print(f"\nSource:\t{args.original_file_name}, {args.original_file_size} bytes\n")

    #        # We can also access the source document from here.
    #        self.assertTrue(args.document.original_file_name.endswith("Rendering.docx"))

    #        self.assertTrue(args.is_export_needed)
    #        self.assertTrue(args.is_subsetting_needed)

    #        # There are two ways of saving an exported font.
    #        # 1 -  Save it to a local file system location:
    #        args.font_file_name = args.original_file_name.split(os.path.sep)[-1]

    #        # 2 -  Save it to a stream:
    #        args.font_stream = open(ARTIFACTS_DIR + args.original_file_name.split(os.path.sep)[-1], "wb")
    #        self.assertFalse(args.keep_font_stream_open)

    ##ExEnd

    def test_html_versions(self):

        for html_version in (aw.saving.HtmlVersion.HTML5,
                             aw.saving.HtmlVersion.XHTML):
            with self.subTest(html_version=html_version):
                #ExStart
                #ExFor:HtmlSaveOptions.__init__(SaveFormat)
                #ExFor:HtmlSaveOptions.html_version
                #ExFor:HtmlVersion
                #ExSummary:Shows how to save a document to a specific version of HTML.
                doc = aw.Document(MY_DIR + "Rendering.docx")

                options = aw.saving.HtmlSaveOptions(aw.SaveFormat.HTML)
                options.html_version = html_version
                options.pretty_format = True

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.html_versions.html", options)

                # Our HTML documents will have minor differences to be compatible with different HTML versions.
                with open(ARTIFACTS_DIR + "HtmlSaveOptions.html_versions.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if html_version == aw.saving.HtmlVersion.HTML5:
                    self.assertIn(
                        "<a id=\"_Toc76372689\"></a>",
                        out_doc_contents)
                    self.assertIn(
                        "<a id=\"_Toc76372689\"></a>",
                        out_doc_contents)
                    self.assertIn(
                        "<table style=\"-aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">",
                        out_doc_contents)

                elif html_version == aw.saving.HtmlVersion.XHTML:
                    self.assertIn(
                        "<a name=\"_Toc76372689\"></a>",
                        out_doc_contents)
                    self.assertIn(
                        "<ul type=\"disc\" style=\"margin:0pt; padding-left:0pt\">",
                        out_doc_contents)
                    self.assertIn(
                        "<table cellspacing=\"0\" cellpadding=\"0\" style=\"-aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\"",
                        out_doc_contents)

                #ExEnd

    def test_export_xhtml_transitional(self):

        for show_doctype_declaration in (False, True):
            with self.subTest(show_doctype_declaration=show_doctype_declaration):
                #ExStart
                #ExFor:HtmlSaveOptions.export_xhtml_transitional
                #ExFor:HtmlSaveOptions.html_version
                #ExFor:HtmlVersion
                #ExSummary:Shows how to display a DOCTYPE heading when converting documents to the Xhtml 1.0 transitional standard.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.writeln("Hello world!")

                options = aw.saving.HtmlSaveOptions(aw.SaveFormat.HTML)
                options.html_version = aw.saving.HtmlVersion.XHTML
                options.export_xhtml_transitional = show_doctype_declaration
                options.pretty_format = True

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.export_xhtml_transitional.html", options)

                # Our document will only contain a DOCTYPE declaration heading if we have set the "export_xhtml_transitional" flag to "True".
                with open(ARTIFACTS_DIR + "HtmlSaveOptions.export_xhtml_transitional.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if show_doctype_declaration:
                    self.assertIn(
                        "<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"no\"?>\n" +
                        "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">\n" +
                        "<html xmlns=\"http://www.w3.org/1999/xhtml\">",
                        out_doc_contents)
                else:
                    self.assertIn("<html>", out_doc_contents)
                #ExEnd

    def test_epub_headings(self):

        #ExStart
        #ExFor:HtmlSaveOptions.epub_navigation_map_level
        #ExSummary:Shows how to filter headings that appear in the navigation panel of a saved Epub document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Every paragraph that we format using a "Heading" style can serve as a heading.
        # Each heading may also have a heading level, determined by the number of its heading style.
        # The headings below are of levels 1-3.
        builder.paragraph_format.style = builder.document.styles.get_by_name("Heading 1")
        builder.writeln("Heading #1")
        builder.paragraph_format.style = builder.document.styles.get_by_name("Heading 2")
        builder.writeln("Heading #2")
        builder.paragraph_format.style = builder.document.styles.get_by_name("Heading 3")
        builder.writeln("Heading #3")
        builder.paragraph_format.style = builder.document.styles.get_by_name("Heading 1")
        builder.writeln("Heading #4")
        builder.paragraph_format.style = builder.document.styles.get_by_name("Heading 2")
        builder.writeln("Heading #5")
        builder.paragraph_format.style = builder.document.styles.get_by_name("Heading 3")
        builder.writeln("Heading #6")

        # Epub readers typically create a table of contents for their documents.
        # Each paragraph with a "Heading" style in the document will create an entry in this table of contents.
        # We can use the "epub_navigation_map_level" property to set a maximum heading level.
        # The Epub reader will not add headings with a level above the one we specify to the contents table.
        options = aw.saving.HtmlSaveOptions(aw.SaveFormat.EPUB)
        options.epub_navigation_map_level = 2

        # Our document has six headings, two of which are above level 2.
        # The table of contents for this document will have four entries.
        doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.epub_headings.epub", options)
        #ExEnd

        self.verify_doc_package_file_contains_string("<navLabel><text>Heading #1</text></navLabel>",
            ARTIFACTS_DIR + "HtmlSaveOptions.epub_headings.epub", "OEBPS/HtmlSaveOptions.epub_headings.ncx")
        self.verify_doc_package_file_contains_string("<navLabel><text>Heading #2</text></navLabel>",
            ARTIFACTS_DIR + "HtmlSaveOptions.epub_headings.epub", "OEBPS/HtmlSaveOptions.epub_headings.ncx")
        self.verify_doc_package_file_contains_string("<navLabel><text>Heading #4</text></navLabel>",
            ARTIFACTS_DIR + "HtmlSaveOptions.epub_headings.epub", "OEBPS/HtmlSaveOptions.epub_headings.ncx")
        self.verify_doc_package_file_contains_string("<navLabel><text>Heading #5</text></navLabel>",
            ARTIFACTS_DIR + "HtmlSaveOptions.epub_headings.epub", "OEBPS/HtmlSaveOptions.epub_headings.ncx")

        with self.assertRaises(Exception):
            self.verify_doc_package_file_contains_string("<navLabel><text>Heading #3</text></navLabel>",
                ARTIFACTS_DIR + "HtmlSaveOptions.epub_headings.epub", "OEBPS/HtmlSaveOptions.epub_headings.ncx")

        with self.assertRaises(Exception):
            self.verify_doc_package_file_contains_string("<navLabel><text>Heading #6</text></navLabel>",
                ARTIFACTS_DIR + "HtmlSaveOptions.epub_headings.epub", "OEBPS/HtmlSaveOptions.epub_headings.ncx")

    def test_doc2_epub_save_options(self):

        #ExStart
        #ExFor:DocumentSplitCriteria
        #ExFor:HtmlSaveOptions
        #ExFor:HtmlSaveOptions.__init__()
        #ExFor:HtmlSaveOptions.encoding
        #ExFor:HtmlSaveOptions.document_split_criteria
        #ExFor:HtmlSaveOptions.export_document_properties
        #ExFor:HtmlSaveOptions.save_format
        #ExFor:SaveOptions
        #ExFor:SaveOptions.save_format
        #ExSummary:Shows how to use a specific encoding when saving a document to .epub.
        doc = aw.Document(MY_DIR + "Rendering.docx")

        # Use a SaveOptions object to specify the encoding for a document that we will save.
        save_options = aw.saving.HtmlSaveOptions()
        save_options.save_format = aw.SaveFormat.EPUB
        save_options.encoding = "utf-8"

        # By default, an output .epub document will have all its contents in one HTML part.
        # A split criterion allows us to segment the document into several HTML parts.
        # We will set the criteria to split the document into heading paragraphs.
        # This is useful for readers who cannot read HTML files more significant than a specific size.
        save_options.document_split_criteria = aw.saving.DocumentSplitCriteria.HEADING_PARAGRAPH

        # Specify that we want to export document properties.
        save_options.export_document_properties = True

        doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.doc2_epub_save_options.epub", save_options)
        #ExEnd

    def test_content_id_urls(self):

        for export_cid_urls_for_mhtml_resources in (False, True):
            with self.subTest(export_cid_urls_for_mhtml_resources=export_cid_urls_for_mhtml_resources):
                #ExStart
                #ExFor:HtmlSaveOptions.export_cid_urls_for_mhtml_resources
                #ExSummary:Shows how to enable content IDs for output MHTML documents.
                doc = aw.Document(MY_DIR + "Rendering.docx")

                # Setting this flag will replace "Content-Location" tags
                # with "Content-ID" tags for each resource from the input document.
                options = aw.saving.HtmlSaveOptions(aw.SaveFormat.MHTML)
                options.export_cid_urls_for_mhtml_resources = export_cid_urls_for_mhtml_resources
                options.css_style_sheet_type = aw.saving.CssStyleSheetType.EXTERNAL
                options.export_font_resources = True
                options.pretty_format = True

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.content_id_urls.mht", options)

                with open(ARTIFACTS_DIR + "HtmlSaveOptions.content_id_urls.mht", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if export_cid_urls_for_mhtml_resources:
                    self.assertIn("Content-ID: <document.html>", out_doc_contents)
                    self.assertIn("<link href=3D\"cid:styles.css\" type=3D\"text/css\" rel=3D\"stylesheet\" />", out_doc_contents)
                    self.assertIn("@font-face { font-family:'Arial Black'; src:url('cid:ariblk.ttf') }", out_doc_contents)
                    self.assertIn("<img src=3D\"cid:image.003.jpeg\" width=3D\"350\" height=3D\"180\" alt=3D\"\" />", out_doc_contents)
                else:
                    self.assertIn("Content-Location: document.html", out_doc_contents)
                    self.assertIn("<link href=3D\"styles.css\" type=3D\"text/css\" rel=3D\"stylesheet\" />", out_doc_contents)
                    self.assertIn("@font-face { font-family:'Arial Black'; src:url('ariblk.ttf') }", out_doc_contents)
                    self.assertIn("<img src=3D\"image.003.jpeg\" width=3D\"350\" height=3D\"180\" alt=3D\"\" />",out_doc_contents)

                #ExEnd

    def test_drop_down_form_field(self):

        for export_drop_down_form_field_as_text in (False, True):
            with self.subTest(export_drop_down_form_field_as_text=export_drop_down_form_field_as_text):
                #ExStart
                #ExFor:HtmlSaveOptions.export_drop_down_form_field_as_text
                #ExSummary:Shows how to get drop-down combo box form fields to blend in with paragraph text when saving to html.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Use a document builder to insert a combo box with the value "Two" selected.
                builder.insert_combo_box("MyComboBox", ["One", "Two", "Three"], 1)

                # The "export_drop_down_form_field_as_text" flag of this SaveOptions object allows us to
                # control how saving the document to HTML treats drop-down combo boxes.
                # Setting it to "True" will convert each combo box into simple text
                # that displays the combo box's currently selected value, effectively freezing it.
                # Setting it to "False" will preserve the functionality of the combo box using <select> and <option> tags.
                options = aw.saving.HtmlSaveOptions()
                options.export_drop_down_form_field_as_text = export_drop_down_form_field_as_text

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.drop_down_form_field.html", options)

                with open(ARTIFACTS_DIR + "HtmlSaveOptions.drop_down_form_field.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if export_drop_down_form_field_as_text:
                    self.assertIn("<span>Two</span>", out_doc_contents)
                else:
                    self.assertIn(
                        "<select name=\"MyComboBox\">" +
                            "<option>One</option>" +
                            "<option selected=\"selected\">Two</option>" +
                            "<option>Three</option>" +
                        "</select>",
                        out_doc_contents)
                #ExEnd

    def test_export_images_as_base64(self):

        for export_items_as_base64 in (False, True):
            with self.subTest(export_items_as_base64=export_items_as_base64):
                #ExStart
                #ExFor:HtmlSaveOptions.export_fonts_as_base64
                #ExFor:HtmlSaveOptions.export_images_as_base64
                #ExSummary:Shows how to save a .html document with images embedded inside it.
                doc = aw.Document(MY_DIR + "Rendering.docx")

                options = aw.saving.HtmlSaveOptions()
                options.export_images_as_base64 = export_items_as_base64
                options.pretty_format = True

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.export_images_as_base64.html", options)

                with open(ARTIFACTS_DIR + "HtmlSaveOptions.export_images_as_base64.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if export_items_as_base64:
                    self.assertIn("<img src=\"data:image/png;base64", out_doc_contents)
                else:
                    self.assertIn("<img src=\"HtmlSaveOptions.export_images_as_base64.001.png\"", out_doc_contents)

                #ExEnd

    def test_export_fonts_as_base64(self):

        #ExStart
        #ExFor:HtmlSaveOptions.export_fonts_as_base64
        #ExFor:HtmlSaveOptions.export_images_as_base64
        #ExSummary:Shows how to embed fonts inside a saved HTML document.
        doc = aw.Document(MY_DIR + "Rendering.docx")

        options = aw.saving.HtmlSaveOptions()
        options.export_fonts_as_base64 = True
        options.css_style_sheet_type = aw.saving.CssStyleSheetType.EMBEDDED
        options.pretty_format = True

        doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.export_fonts_as_base64.html", options)
        #ExEnd

    def test_export_language_information(self):

        for export_language_information in (False, True):
            with self.subTest(export_language_information=export_language_information):
                #ExStart
                #ExFor:HtmlSaveOptions.export_language_information
                #ExSummary:Shows how to preserve language information when saving to .html.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Use the builder to write text while formatting it in different locales.
                builder.font.locale_id = 1033 # en-US
                builder.writeln("Hello world!")

                builder.font.locale_id = 2057 # en-GB
                builder.writeln("Hello again!")

                builder.font.locale_id = 1049 # ru-RU
                builder.write("Привет, мир!")

                # When saving the document to HTML, we can pass a SaveOptions object
                # to either preserve or discard each formatted text's locale.
                # If we set the "export_language_information" flag to "True",
                # the output HTML document will contain the locales in "lang" attributes of <span> tags.
                # If we set the "export_language_information" flag to "False',
                # the text in the output HTML document will not contain any locale information.
                options = aw.saving.HtmlSaveOptions()
                options.export_language_information = export_language_information
                options.pretty_format = True

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.export_language_information.html", options)

                with open(ARTIFACTS_DIR + "HtmlSaveOptions.export_language_information.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if export_language_information:
                    self.assertIn("<span>Hello world!</span>", out_doc_contents)
                    self.assertIn("<span lang=\"en-GB\">Hello again!</span>", out_doc_contents)
                    self.assertIn("<span lang=\"ru-RU\">Привет, мир!</span>", out_doc_contents)
                else:
                    self.assertIn("<span>Hello world!</span>", out_doc_contents)
                    self.assertIn("<span>Hello again!</span>", out_doc_contents)
                    self.assertIn("<span>Привет, мир!</span>", out_doc_contents)

                #ExEnd

    def test_list(self):

        for export_list_labels in (aw.saving.ExportListLabels.AS_INLINE_TEXT,
                                   aw.saving.ExportListLabels.AUTO,
                                   aw.saving.ExportListLabels.BY_HTML_TAGS):
            with self.subTest(export_list_labels=export_list_labels):
                #ExStart
                #ExFor:ExportListLabels
                #ExFor:HtmlSaveOptions.export_list_labels
                #ExSummary:Shows how to configure list exporting to HTML.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                list = doc.lists.add(aw.lists.ListTemplate.NUMBER_DEFAULT)
                builder.list_format.list = list

                builder.writeln("Default numbered list item 1.")
                builder.writeln("Default numbered list item 2.")
                builder.list_format.list_indent()
                builder.writeln("Default numbered list item 3.")
                builder.list_format.remove_numbers()

                list = doc.lists.add(aw.lists.ListTemplate.OUTLINE_HEADINGS_LEGAL)
                builder.list_format.list = list

                builder.writeln("Outline legal heading list item 1.")
                builder.writeln("Outline legal heading list item 2.")
                builder.list_format.list_indent()
                builder.writeln("Outline legal heading list item 3.")
                builder.list_format.list_indent()
                builder.writeln("Outline legal heading list item 4.")
                builder.list_format.list_indent()
                builder.writeln("Outline legal heading list item 5.")
                builder.list_format.remove_numbers()

                # When saving the document to HTML, we can pass a SaveOptions object
                # to decide which HTML elements the document will use to represent lists.
                # Setting the "export_list_labels" property to "ExportListLabels.AS_INLINE_TEXT"
                # will create lists by formatting spans.
                # Setting the "export_list_labels" property to "ExportListLabels.AUTO" will use the <p> tag
                # to build lists in cases when using the <ol> and <li> tags may cause loss of formatting.
                # Setting the "export_list_labels" property to "ExportListLabels.BY_HTML_TAGS"
                # will use <ol> and <li> tags to build all lists.
                options = aw.saving.HtmlSaveOptions()
                options.export_list_labels = export_list_labels

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.list.html", options)

                with open(ARTIFACTS_DIR + "HtmlSaveOptions.list.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if export_list_labels == aw.saving.ExportListLabels.AS_INLINE_TEXT:
                    self.assertIn(
                        "<p style=\"margin-top:0pt; margin-left:72pt; margin-bottom:0pt; text-indent:-18pt; -aw-import:list-item; -aw-list-level-number:1; -aw-list-number-format:'%1.'; -aw-list-number-styles:'lowerLetter'; -aw-list-number-values:'1'; -aw-list-padding-sml:9.67pt\">" +
                            "<span style=\"-aw-import:ignore\">" +
                                "<span>a.</span>" +
                                "<span style=\"width:9.67pt; font:7pt 'Times New Roman'; display:inline-block; -aw-import:spaces\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" +
                            "</span>" +
                            "<span>Default numbered list item 3.</span>" +
                        "</p>", out_doc_contents)

                    self.assertIn(
                        "<p style=\"margin-top:0pt; margin-left:43.2pt; margin-bottom:0pt; text-indent:-43.2pt; -aw-import:list-item; -aw-list-level-number:3; -aw-list-number-format:'%0.%1.%2.%3'; -aw-list-number-styles:'decimal decimal decimal decimal'; -aw-list-number-values:'2 1 1 1'; -aw-list-padding-sml:10.2pt\">" +
                            "<span style=\"-aw-import:ignore\">" +
                                "<span>2.1.1.1</span>" +
                                "<span style=\"width:10.2pt; font:7pt 'Times New Roman'; display:inline-block; -aw-import:spaces\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" +
                            "</span>" +
                            "<span>Outline legal heading list item 5.</span>" +
                        "</p>", out_doc_contents)

                elif export_list_labels == aw.saving.ExportListLabels.AUTO:
                    self.assertIn(
                        "<ol type=\"a\" style=\"margin-right:0pt; margin-left:0pt; padding-left:0pt\">" +
                            "<li style=\"margin-left:31.33pt; padding-left:4.67pt\">" +
                                "<span>Default numbered list item 3.</span>" +
                            "</li>" +
                        "</ol>", out_doc_contents)

                    self.assertIn(
                        "<p style=\"margin-top:0pt; margin-left:43.2pt; margin-bottom:0pt; text-indent:-43.2pt; -aw-import:list-item; -aw-list-level-number:3; " +
                        "-aw-list-number-format:'%0.%1.%2.%3'; -aw-list-number-styles:'decimal decimal decimal decimal'; " +
                        "-aw-list-number-values:'2 1 1 1'; -aw-list-padding-sml:10.2pt\">" +
                            "<span style=\"-aw-import:ignore\">" +
                                "<span>2.1.1.1</span>" +
                                "<span style=\"width:10.2pt; font:7pt 'Times New Roman'; display:inline-block; -aw-import:spaces\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" +
                            "</span>" +
                            "<span>Outline legal heading list item 5.</span>" +
                        "</p>", out_doc_contents)

                elif export_list_labels == aw.saving.ExportListLabels.BY_HTML_TAGS:
                    self.assertIn(
                        "<ol type=\"a\" style=\"margin-right:0pt; margin-left:0pt; padding-left:0pt\">" +
                            "<li style=\"margin-left:31.33pt; padding-left:4.67pt\">" +
                                "<span>Default numbered list item 3.</span>" +
                            "</li>" +
                        "</ol>", out_doc_contents)

                    self.assertIn(
                        "<ol type=\"1\" class=\"awlist3\" style=\"margin-right:0pt; margin-left:0pt; padding-left:0pt\">" +
                            "<li style=\"margin-left:7.2pt; text-indent:-43.2pt; -aw-list-padding-sml:10.2pt\">" +
                                "<span style=\"width:10.2pt; font:7pt 'Times New Roman'; display:inline-block; -aw-import:ignore\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" +
                                "<span>Outline legal heading list item 5.</span>" +
                            "</li>" +
                        "</ol>", out_doc_contents)

                #ExEnd

    def test_export_page_margins(self):

        for export_page_margins in (False, True):
            with self.subTest(export_page_margins=export_page_margins):
                #ExStart
                #ExFor:HtmlSaveOptions.export_page_margins
                #ExSummary:Shows how to show out-of-bounds objects in output HTML documents.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Use a builder to insert a shape with no wrapping.
                shape = builder.insert_shape(aw.drawing.ShapeType.CUBE, 200, 200)

                shape.relative_horizontal_position = aw.drawing.RelativeHorizontalPosition.PAGE
                shape.relative_vertical_position = aw.drawing.RelativeVerticalPosition.PAGE
                shape.wrap_type = aw.drawing.WrapType.NONE

                # Negative shape position values may place the shape outside of page boundaries.
                # If we export this to HTML, the shape will appear truncated.
                shape.left = -150

                # When saving the document to HTML, we can pass a SaveOptions object
                # to decide whether to adjust the page to display out-of-bounds objects fully.
                # If we set the "export_page_margins" flag to "True", the shape will be fully visible in the output HTML.
                # If we set the "export_page_margins" flag to "False",
                # our document will display the shape truncated as we would see it in Microsoft Word.
                options = aw.saving.HtmlSaveOptions()
                options.export_page_margins = export_page_margins

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.export_page_margins.html", options)

                with open(ARTIFACTS_DIR + "HtmlSaveOptions.export_page_margins.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if export_page_margins:
                    self.assertIn("<style type=\"text/css\">div.Section1 { margin:70.85pt }</style>", out_doc_contents)
                    self.assertIn("<div class=\"Section1\"><p style=\"margin-top:0pt; margin-left:151pt; margin-bottom:0pt\">", out_doc_contents)
                else:
                    self.assertNotIn("style type=\"text/css\">", out_doc_contents)
                    self.assertIn("<div><p style=\"margin-top:0pt; margin-left:221.85pt; margin-bottom:0pt\">", out_doc_contents)

                #ExEnd

    def test_export_page_setup(self):

        for export_page_setup in (False, True):
            with self.subTest(export_page_setup=export_page_setup):
                #ExStart
                #ExFor:HtmlSaveOptions.export_page_setup
                #ExSummary:Shows how decide whether to preserve section structure/page setup information when saving to HTML.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.write("Section 1")
                builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
                builder.write("Section 2")

                page_setup = doc.sections[0].page_setup
                page_setup.top_margin = 36.0
                page_setup.bottom_margin = 36.0
                page_setup.paper_size = aw.PaperSize.A5

                # When saving the document to HTML, we can pass a SaveOptions object
                # to decide whether to preserve or discard page setup settings.
                # If we set the "export_page_setup" flag to "True", the output HTML document will contain our page setup configuration.
                # If we set the "export_page_setup" flag to "False", the save operation will discard our page setup settings
                # for the first section, and both sections will look identical.
                options = aw.saving.HtmlSaveOptions()
                options.export_page_setup = export_page_setup

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.export_page_setup.html", options)

                with open(ARTIFACTS_DIR + "HtmlSaveOptions.export_page_setup.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if export_page_setup:
                    self.assertIn(
                        "<style type=\"text/css\">" +
                            "@page Section1 { size:419.55pt 595.3pt; margin:36pt 70.85pt }" +
                            "@page Section2 { size:612pt 792pt; margin:70.85pt }" +
                            "div.Section1 { page:Section1 }div.Section2 { page:Section2 }" +
                        "</style>", out_doc_contents)

                    self.assertIn(
                        "<div class=\"Section1\">" +
                            "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                                "<span>Section 1</span>" +
                            "</p>" +
                        "</div>", out_doc_contents)
                else:
                    self.assertNotIn("style type=\"text/css\">", out_doc_contents)

                    self.assertIn(
                        "<div>" +
                            "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                                "<span>Section 1</span>" +
                            "</p>" +
                        "</div>", out_doc_contents)

                #ExEnd

    def test_relative_font_size(self):

        for export_relative_font_size in (False, True):
            with self.subTest(export_relative_font_size=export_relative_font_size):
                #ExStart
                #ExFor:HtmlSaveOptions.export_relative_font_size
                #ExSummary:Shows how to use relative font sizes when saving to .html.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.writeln("Default font size, ")
                builder.font.size = 24
                builder.writeln("2x default font size,")
                builder.font.size = 96
                builder.write("8x default font size")

                # When we save the document to HTML, we can pass a SaveOptions object
                # to determine whether to use relative or absolute font sizes.
                # Set the "export_relative_font_size" flag to "True" to declare font sizes
                # using the "em" measurement unit, which is a factor that multiplies the current font size.
                # Set the "export_relative_font_size" flag to "False" to declare font sizes
                # using the "pt" measurement unit, which is the font's absolute size in points.
                options = aw.saving.HtmlSaveOptions()
                options.export_relative_font_size = export_relative_font_size

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.relative_font_size.html", options)

                with open(ARTIFACTS_DIR + "HtmlSaveOptions.relative_font_size.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if export_relative_font_size:
                    self.assertIn(
                        "<body style=\"font-family:'Times New Roman'\">" +
                            "<div>" +
                                "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                                    "<span>Default font size, </span>" +
                                "</p>" +
                                "<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:2em\">" +
                                    "<span>2x default font size,</span>" +
                                "</p>" +
                                "<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:8em\">" +
                                    "<span>8x default font size</span>" +
                                "</p>" +
                            "</div>" +
                        "</body>", out_doc_contents)
                else:
                    self.assertIn(
                        "<body style=\"font-family:'Times New Roman'; font-size:12pt\">" +
                            "<div>" +
                                "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                                    "<span>Default font size, </span>" +
                                "</p>" +
                                "<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:24pt\">" +
                                    "<span>2x default font size,</span>" +
                                "</p>" +
                                "<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:96pt\">" +
                                    "<span>8x default font size</span>" +
                                "</p>" +
                            "</div>" +
                        "</body>", out_doc_contents)

                #ExEnd

    def test_export_text_box(self):

        for export_text_box_as_svg in (False, True):
            with self.subTest(export_text_box_as_svg=export_text_box_as_svg):
                #ExStart
                #ExFor:HtmlSaveOptions.export_text_box_as_svg
                #ExSummary:Shows how to export text boxes as scalable vector graphics.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                text_box = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 100.0, 60.0)
                builder.move_to(text_box.first_paragraph)
                builder.write("My text box")

                # When we save the document to HTML, we can pass a SaveOptions object
                # to determine how the saving operation will export text box shapes.
                # If we set the "export_text_box_as_svg" flag to "True",
                # the save operation will convert shapes with text into SVG objects.
                # If we set the "export_text_box_as_svg" flag to "False",
                # the save operation will convert shapes with text into images.
                options = aw.saving.HtmlSaveOptions()
                options.export_text_box_as_svg = export_text_box_as_svg

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.export_text_box.html", options)

                with open(ARTIFACTS_DIR + "HtmlSaveOptions.export_text_box.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if export_text_box_as_svg:
                    self.assertIn(
                        "<span style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\">" +
                        "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"133\" height=\"80\">",
                        out_doc_contents)
                else:
                    self.assertIn(
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<img src=\"HtmlSaveOptions.export_text_box.001.png\" width=\"136\" height=\"83\" alt=\"\" " +
                            "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                        "</p>", out_doc_contents)

                #ExEnd

    def test_round_trip_information(self):

        for export_roundtrip_information in (False, True):
            with self.subTest(export_roundtrip_information=export_roundtrip_information):
                #ExStart
                #ExFor:HtmlSaveOptions.export_roundtrip_information
                #ExSummary:Shows how to preserve hidden elements when converting to .html.
                doc = aw.Document(MY_DIR + "Rendering.docx")

                # When converting a document to .html, some elements such as hidden bookmarks, original shape positions,
                # or footnotes will be either removed or converted to plain text and effectively be lost.
                # Saving with a HtmlSaveOptions object with "export_roundtrip_information" set to True will preserve these elements.

                # When we save the document to HTML, we can pass a SaveOptions object to determine
                # how the saving operation will export document elements that HTML does not support or use,
                # such as hidden bookmarks and original shape positions.
                # If we set the "export_roundtrip_information" flag to "True", the save operation will preserve these elements.
                # If we set the "export_roundtrip_information" flag to "False", the save operation will discard these elements.
                # We will want to preserve such elements if we intend to load the saved HTML using Aspose.Words,
                # as they could be of use once again.
                options = aw.saving.HtmlSaveOptions()
                options.export_roundtrip_information = export_roundtrip_information

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.round_trip_information.html", options)

                with open(ARTIFACTS_DIR + "HtmlSaveOptions.round_trip_information.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                doc = aw.Document(ARTIFACTS_DIR + "HtmlSaveOptions.round_trip_information.html")

                if export_roundtrip_information:
                    self.assertIn("<div style=\"-aw-headerfooter-type:header-primary; clear:both\">", out_doc_contents)
                    self.assertIn("<span style=\"-aw-import:ignore\">&#xa0;</span>", out_doc_contents)

                    self.assertIn(
                        "td colspan=\"2\" style=\"width:210.6pt; border-style:solid; border-width:0.75pt 6pt 0.75pt 0.75pt; " +
                        "padding-right:2.4pt; padding-left:5.03pt; vertical-align:top; " +
                        "-aw-border-bottom:0.5pt single; -aw-border-left:0.5pt single; -aw-border-top:0.5pt single\">",
                        out_doc_contents)

                    self.assertIn(
                        "<li style=\"margin-left:30.2pt; padding-left:5.8pt; -aw-font-family:'Courier New'; -aw-font-weight:normal; -aw-number-format:'o'\">",
                        out_doc_contents)

                    self.assertIn(
                        "<img src=\"HtmlSaveOptions.round_trip_information.003.jpeg\" width=\"350\" height=\"180\" alt=\"\" " +
                        "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />",
                        out_doc_contents)

                    self.assertIn(
                        "<span>Page number </span>" +
                        "<span style=\"-aw-field-start:true\"></span>" +
                        "<span style=\"-aw-field-code:' PAGE   \\\\* MERGEFORMAT '\"></span>" +
                        "<span style=\"-aw-field-separator:true\"></span>" +
                        "<span>1</span>" +
                        "<span style=\"-aw-field-end:true\"></span>",
                        out_doc_contents)

                    self.assertEqual(1, len([f for f in doc.range.fields if f.type == aw.fields.FieldType.FIELD_PAGE]))
                else:
                    self.assertIn("<div style=\"clear:both\">", out_doc_contents)
                    self.assertIn("<span>&#xa0;</span>", out_doc_contents)

                    self.assertIn(
                        "<td colspan=\"2\" style=\"width:210.6pt; border-style:solid; border-width:0.75pt 6pt 0.75pt 0.75pt; " +
                        "padding-right:2.4pt; padding-left:5.03pt; vertical-align:top\">",
                        out_doc_contents)

                    self.assertIn(
                        "<li style=\"margin-left:30.2pt; padding-left:5.8pt\">",
                        out_doc_contents)

                    self.assertIn(
                        "<img src=\"HtmlSaveOptions.round_trip_information.003.jpeg\" width=\"350\" height=\"180\" alt=\"\" />",
                        out_doc_contents)

                    self.assertIn("<span>Page number 1</span>", out_doc_contents)

                    self.assertEqual(0, len([f for f in doc.range.fields if f.type == aw.fields.FieldType.FIELD_PAGE]))

                #ExEnd

    def test_export_toc_page_numbers(self):

        for export_toc_page_numbers in (False, True):
            with self.subTest(export_toc_page_numbers=export_toc_page_numbers):
                #ExStart
                #ExFor:HtmlSaveOptions.export_toc_page_numbers
                #ExSummary:Shows how to display page numbers when saving a document with a table of contents to .html.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Insert a table of contents, and then populate the document with paragraphs formatted using a "Heading"
                # style that the table of contents will pick up as entries. Each entry will display the heading paragraph on the left,
                # and the page number that contains the heading on the right.
                field_toc = builder.insert_field(aw.fields.FieldType.FIELD_TOC, True).as_field_toc()

                builder.paragraph_format.style = builder.document.styles.get_by_name("Heading 1")
                builder.insert_break(aw.BreakType.PAGE_BREAK)
                builder.writeln("Entry 1")
                builder.writeln("Entry 2")
                builder.insert_break(aw.BreakType.PAGE_BREAK)
                builder.writeln("Entry 3")
                builder.insert_break(aw.BreakType.PAGE_BREAK)
                builder.writeln("Entry 4")
                field_toc.update_page_numbers()
                doc.update_fields()

                # HTML documents do not have pages. If we save this document to HTML,
                # the page numbers that our TOC displays will have no meaning.
                # When we save the document to HTML, we can pass a SaveOptions object to omit these page numbers from the TOC.
                # If we set the "export_toc_page_numbers" flag to "True",
                # each TOC entry will display the heading, separator, and page number, preserving its appearance in Microsoft Word.
                # If we set the "export_toc_page_numbers" flag to "False",
                # the save operation will omit both the separator and page number and leave the heading for each entry intact.
                options = aw.saving.HtmlSaveOptions()
                options.export_toc_page_numbers = export_toc_page_numbers

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.export_toc_page_numbers.html", options)

                with open(ARTIFACTS_DIR + "HtmlSaveOptions.export_toc_page_numbers.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if export_toc_page_numbers:
                    self.assertIn(
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                        "<span>Entry 1</span>" +
                        "<span style=\"width:428.14pt; font-family:'Lucida Console'; font-size:10pt; display:inline-block; -aw-font-family:'Times New Roman'; " +
                        "-aw-tabstop-align:right; -aw-tabstop-leader:dots; -aw-tabstop-pos:469.8pt\">.......................................................................</span>" +
                        "<span>2</span>" +
                        "</p>", out_doc_contents)
                else:
                    self.assertIn(
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                        "<span>Entry 1</span>" +
                        "</p>", out_doc_contents)

                #ExEnd

    def test_font_subsetting(self):

        for font_resources_subsetting_size_threshold in (0, 1000000, 2**31 - 1):
            with self.subTest(font_resources_subsetting_size_threshold=font_resources_subsetting_size_threshold):
                #ExStart
                #ExFor:HtmlSaveOptions.font_resources_subsetting_size_threshold
                #ExSummary:Shows how to work with font subsetting.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.font.name = "Arial"
                builder.writeln("Hello world!")
                builder.font.name = "Times New Roman"
                builder.writeln("Hello world!")
                builder.font.name = "Courier New"
                builder.writeln("Hello world!")

                # When we save the document to HTML, we can pass a SaveOptions object configure font subsetting.
                # Suppose we set the "export_font_resources" flag to "True" and also name a folder in the "fonts_folder" property.
                # In that case, the saving operation will create that folder and place a .ttf file inside
                # that folder for each font that our document uses.
                # Each .ttf file will contain that font's entire glyph set,
                # which may potentially result in a very large file that accompanies the document.
                # When we apply subsetting to a font, its exported raw data will only contain the glyphs that the document is
                # using instead of the entire glyph set. If the text in our document only uses a small fraction of a font's
                # glyph set, then subsetting will significantly reduce our output documents' size.
                # We can use the "font_resources_subsetting_size_threshold" property to define a .ttf file size, in bytes.
                # If an exported font creates a size bigger file than that, then the save operation will apply subsetting to that font.
                # Setting a threshold of 0 applies subsetting to all fonts,
                # and setting it to "2**31 - 1" effectively disables subsetting.
                fonts_folder = ARTIFACTS_DIR + "HtmlSaveOptions.font_subsetting.fonts"
                
                if os.path.exists(fonts_folder):
                    shutil.rmtree(fonts_folder)

                options = aw.saving.HtmlSaveOptions()
                options.export_font_resources = True
                options.fonts_folder = fonts_folder
                options.font_resources_subsetting_size_threshold = font_resources_subsetting_size_threshold

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.font_subsetting.html", options)

                font_file_names = glob.glob(fonts_folder + "/*.ttf")

                self.assertEqual(3, len(font_file_names))

                for filename in font_file_names:
                    # By default, the .ttf files for each of our three fonts will be over 700MB.
                    # Subsetting will reduce them all to under 30MB.
                    font_file_size = os.path.getsize(filename)

                    self.assertTrue(font_file_size > 700000 or font_file_size < 30000)
                    self.assertTrue(max(font_resources_subsetting_size_threshold, 30000) > font_file_size)

                #ExEnd

    def test_metafile_format(self):

        for html_metafile_format in (aw.saving.HtmlMetafileFormat.PNG,
                                     aw.saving.HtmlMetafileFormat.SVG,
                                     aw.saving.HtmlMetafileFormat.EMF_OR_WMF):
            with self.subTest(html_metafile_format=html_metafile_format):
                #ExStart
                #ExFor:HtmlMetafileFormat
                #ExFor:HtmlSaveOptions.metafile_format
                #ExFor:HtmlLoadOptions.convert_svg_to_emf
                #ExSummary:Shows how to convert SVG objects to a different format when saving HTML documents.
                html = """
                    <html>
                        <svg xmlns='http://www.w3.org/2000/svg' width='500' height='40' viewBox='0 0 500 40'>
                            <text x='0' y='35' font-family='Verdana' font-size='35'>Hello world!</text>
                        </svg>
                    </html>
                    """

                # Use 'convert_svg_to_emf' to turn back the legacy behavior
                # where all SVG images loaded from an HTML document were converted to EMF.
                # Now SVG images are loaded without conversion
                # if the MS Word version specified in load options supports SVG images natively.
                load_options = aw.loading.HtmlLoadOptions()
                load_options.convert_svg_to_emf = True

                doc = aw.Document(io.BytesIO(html.encode('utf-8')), load_options)

                # This document contains a <svg> element in the form of text.
                # When we save the document to HTML, we can pass a SaveOptions object
                # to determine how the saving operation handles this object.
                # Setting the "metafile_format" property to "HtmlMetafileFormat.PNG" to convert it to a PNG image.
                # Setting the "metafile_format" property to "HtmlMetafileFormat.SVG" preserve it as a SVG object.
                # Setting the "metafile_format" property to "HtmlMetafileFormat.EMF_OR_WMF" to convert it to a metafile.
                options = aw.saving.HtmlSaveOptions()
                options.metafile_format = html_metafile_format

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.metafile_format.html", options)

                with open(ARTIFACTS_DIR + "HtmlSaveOptions.metafile_format.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if html_metafile_format == aw.saving.HtmlMetafileFormat.PNG:
                    self.assertIn(
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<img src=\"HtmlSaveOptions.metafile_format.001.png\" width=\"500\" height=\"40\" alt=\"\" " +
                            "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                        "</p>", out_doc_contents)

                elif html_metafile_format == aw.saving.HtmlMetafileFormat.SVG:
                    self.assertIn(
                        "<span style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\">" +
                        "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"499\" height=\"40\">",
                        out_doc_contents)

                elif html_metafile_format == aw.saving.HtmlMetafileFormat.EMF_OR_WMF:
                    self.assertIn(
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<img src=\"HtmlSaveOptions.metafile_format.001.emf\" width=\"500\" height=\"40\" alt=\"\" " +
                            "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                        "</p>", out_doc_contents)

                #ExEnd

    def test_office_math_output_mode(self):

        for html_office_math_output_mode in (aw.saving.HtmlOfficeMathOutputMode.IMAGE,
                                             aw.saving.HtmlOfficeMathOutputMode.MATH_ML,
                                             aw.saving.HtmlOfficeMathOutputMode.TEXT):
            with self.subTest(html_office_math_output_mode=html_office_math_output_mode):
                #ExStart
                #ExFor:HtmlOfficeMathOutputMode
                #ExFor:HtmlSaveOptions.office_math_output_mode
                #ExSummary:Shows how to specify how to export Microsoft OfficeMath objects to HTML.
                doc = aw.Document(MY_DIR + "Office math.docx")

                # When we save the document to HTML, we can pass a SaveOptions object
                # to determine how the saving operation handles OfficeMath objects.
                # Setting the "office_math_output_mode" property to "HtmlOfficeMathOutputMode.IMAGE"
                # will render each OfficeMath object into an image.
                # Setting the "office_math_output_mode" property to "HtmlOfficeMathOutputMode.MATH_ML"
                # will convert each OfficeMath object into MathML.
                # Setting the "office_math_output_mode" property to "HtmlOfficeMathOutputMode.TEXT"
                # will represent each OfficeMath formula using plain HTML text.
                options = aw.saving.HtmlSaveOptions()
                options.office_math_output_mode = html_office_math_output_mode

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.office_math_output_mode.html", options)

                with open(ARTIFACTS_DIR + "HtmlSaveOptions.office_math_output_mode.html", "rt", encoding="utf-8") as file:
                    out_doc_contents = file.read()

                if html_office_math_output_mode == aw.saving.HtmlOfficeMathOutputMode.IMAGE:
                    self.assertRegex(
                        out_doc_contents,
                        "<p style=\"margin-top:0pt; margin-bottom:10pt\">" +
                            "<img src=\"HtmlSaveOptions.office_math_output_mode.001.png\" width=\"159\" height=\"19\" alt=\"\" style=\"vertical-align:middle; " +
                            "-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                        "</p>")

                elif html_office_math_output_mode == aw.saving.HtmlOfficeMathOutputMode.MATH_ML:
                    self.assertRegex(
                        out_doc_contents,
                        "<p style=\"margin-top:0pt; margin-bottom:10pt; text-align:center\">" +
                            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">" +
                                "<mi>i</mi>" +
                                "<mo>[+]</mo>" +
                                "<mi>b</mi>" +
                                "<mo>-</mo>" +
                                "<mi>c</mi>" +
                                "<mo>≥</mo>" +
                                ".*" +
                            "</math>" +
                        "</p>")

                elif html_office_math_output_mode == aw.saving.HtmlOfficeMathOutputMode.TEXT:
                    self.assertRegex(
                        out_doc_contents,
                        r'<p style="margin-top:0pt; margin-bottom:10pt; text-align:center">' +
                            r'<span style="font-family:\'Cambria Math\'">i[+]b-c≥iM[+]bM-cM </span>' +
                        r'</p>')

                #ExEnd

    def test_scale_image_to_shape_size(self):

        for scale_image_to_shape_size in (False, True):
            with self.subTest(scale_image_to_shape_size=scale_image_to_shape_size):
                #ExStart
                #ExFor:HtmlSaveOptions.scale_image_to_shape_size
                #ExSummary:Shows how to disable the scaling of images to their parent shape dimensions when saving to .html.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Insert a shape which contains an image, and then make that shape considerably smaller than the image.
                image = drawing.Image.from_file(IMAGE_DIR + "Transparent background logo.png")

                self.assertEqual(400, image.size.width)
                self.assertEqual(400, image.size.height)

                image_shape = builder.insert_image(image)
                image_shape.width = 50
                image_shape.height = 50

                # Saving a document that contains shapes with images to HTML will create an image file in the local file system
                # for each such shape. The output HTML document will use <image> tags to link to and display these images.
                # When we save the document to HTML, we can pass a SaveOptions object to determine
                # whether to scale all images that are inside shapes to the sizes of their shapes.
                # Setting the "scale_image_to_shape_size" flag to "True" will shrink every image
                # to the size of the shape that contains it, so that no saved images will be larger than the document requires them to be.
                # Setting the "scale_image_to_shape_size" flag to "False" will preserve these images' original sizes,
                # which will take up more space in exchange for preserving image quality.
                options = aw.saving.HtmlSaveOptions()
                options.scale_image_to_shape_size = scale_image_to_shape_size

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.scale_image_to_shape_size.html", options)

                file_size = os.path.getsize(ARTIFACTS_DIR + "HtmlSaveOptions.scale_image_to_shape_size.001.png")

                if scale_image_to_shape_size:
                    self.assertGreater(10000, file_size)
                else:
                    self.assertLess(30000, file_size)

                #ExEnd

    def test_image_folder(self):

        #ExStart
        #ExFor:HtmlSaveOptions
        #ExFor:HtmlSaveOptions.export_text_input_form_field_as_text
        #ExFor:HtmlSaveOptions.images_folder
        #ExSummary:Shows how to specify the folder for storing linked images after saving to .html.
        doc = aw.Document(MY_DIR + "Rendering.docx")

        images_dir = os.path.join(ARTIFACTS_DIR, "SaveHtmlWithOptions")

        if os.path.exists(images_dir):
            shutil.rmtree(images_dir)

        os.makedirs(images_dir)

        # Set an option to export form fields as plain text instead of HTML input elements.
        options = aw.saving.HtmlSaveOptions(aw.SaveFormat.HTML)
        options.export_text_input_form_field_as_text = True
        options.images_folder = images_dir

        doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.image_folder.html", options)
        #ExEnd

        self.assertTrue(os.path.exists(ARTIFACTS_DIR + "HtmlSaveOptions.image_folder.html"))
        self.assertEqual(9, len(os.listdir(images_dir)))

        shutil.rmtree(images_dir)

    ##ExStart
    ##ExFor:ImageSavingArgs.current_shape
    ##ExFor:ImageSavingArgs.document
    ##ExFor:ImageSavingArgs.image_stream
    ##ExFor:ImageSavingArgs.is_image_available
    ##ExFor:ImageSavingArgs.keep_image_stream_open
    ##ExSummary:Shows how to involve an image saving callback in an HTML conversion process.
    #def test_image_saving_callback(self):

    #    doc = aw.Document(MY_DIR + "Rendering.docx")

    #    # When we save the document to HTML, we can pass a SaveOptions object to designate a callback
    #    # to customize the image saving process.
    #    options = aw.saving.HtmlSaveOptions()
    #    options.image_saving_callback = ExHtmlSaveOptions.ImageShapePrinter()

    #    doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.image_saving_callback.html", options)

    #class ImageShapePrinter(aw.savging.IImageSavingCallback):
    #    """Prints the properties of each image as the saving process saves it to an image file in the local file system
    #    during the exporting of a document to HTML."""

    #    def __init__(self):
    #        self.image_count = 0

    #    def image_saving(self, args: aw.saving.ImageSavingArgs):

    #        args.keep_image_stream_open = False
    #        self.assertTrue(args.is_image_available)

    #        self.image_count += 1
    #        print(f"{args.document.original_file_name.split('\\')[-1]} Image #{self.image_count}")

    #        layout_collector = aw.layout.LayoutCollector(args.document)

    #        print(f"\tOn page:\t{layoutCollector.get_start_page_index(args.current_shape)}")
    #        print(f"\tDimensions:\t{args.current_shape.bounds}")
    #        print(f"\tAlignment:\t{args.current_shape.vertical_alignment}")
    #        print(f"\tWrap type:\t{args.current_shape.wrap_type}")
    #        print(f"Output filename:\t{args.image_file_name}\n")

    ##ExEnd

    def test_pretty_format(self):

        for use_pretty_format in (False, True):
            with self.subTest(use_pretty_format=use_pretty_format):
                #ExStart
                #ExFor:SaveOptions.pretty_format
                #ExSummary:Shows how to enhance the readability of the raw code of a saved .html document.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.writeln("Hello world!")

                html_options = aw.saving.HtmlSaveOptions(aw.SaveFormat.HTML)
                html_options.pretty_format = use_pretty_format

                doc.save(ARTIFACTS_DIR + "HtmlSaveOptions.pretty_format.html", html_options)

                # Enabling pretty format makes the raw html code more readable by adding tab stop and new line characters.
                with open(ARTIFACTS_DIR + "HtmlSaveOptions.pretty_format.html", "rt", encoding="utf-8") as file:
                    html = file.read()

                if use_pretty_format:
                    self.assertEqual(
                        "<html>\n" +
                            "\t<head>\n" +
                                "\t\t<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />\n" +
                                "\t\t<meta http-equiv=\"Content-Style-Type\" content=\"text/css\" />\n" +
                                f"\t\t<meta name=\"generator\" content=\"{aw.BuildVersionInfo.product} {aw.BuildVersionInfo.version}\" />\n" +
                                "\t\t<title>\n" +
                                "\t\t</title>\n" +
                            "\t</head>\n" +
                            "\t<body style=\"font-family:'Times New Roman'; font-size:12pt\">\n" +
                                "\t\t<div>\n" +
                                    "\t\t\t<p style=\"margin-top:0pt; margin-bottom:0pt\">\n" +
                                        "\t\t\t\t<span>Hello world!</span>\n" +
                                    "\t\t\t</p>\n" +
                                    "\t\t\t<p style=\"margin-top:0pt; margin-bottom:0pt\">\n" +
                                        "\t\t\t\t<span style=\"-aw-import:ignore\">&#xa0;</span>\n" +
                                    "\t\t\t</p>\n" +
                                "\t\t</div>\n" +
                            "\t</body>\n</html>",
                        html)
                else:
                    self.assertEqual(
                        "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />" +
                        "<meta http-equiv=\"Content-Style-Type\" content=\"text/css\" />" +
                        f"<meta name=\"generator\" content=\"{aw.BuildVersionInfo.product} {aw.BuildVersionInfo.version}\" /><title></title></head>" +
                        "<body style=\"font-family:'Times New Roman'; font-size:12pt\">" +
                        "<div><p style=\"margin-top:0pt; margin-bottom:0pt\"><span>Hello world!</span></p>" +
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\"><span style=\"-aw-import:ignore\">&#xa0;</span></p></div></body></html>",
                        html)
                #ExEnd
