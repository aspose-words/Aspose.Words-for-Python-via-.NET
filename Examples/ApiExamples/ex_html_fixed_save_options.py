import shutil
import glob
import os
from document_helper import DocumentHelper
import sys
# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
import aspose.words.saving
import document_helper
import pathlib
import system_helper
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExHtmlFixedSaveOptions(ApiExampleBase):

    @unittest.skipIf(sys.platform.startswith('win'), 'Discrepancy in assertion between Python and .Net')
    def test_use_encoding(self):
        #ExStart
        #ExFor:HtmlFixedSaveOptions.encoding
        #ExSummary:Shows how to set which encoding to use while exporting a document to HTML.
        from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, GOLDS_DIR, TEMP_DIR, IMAGE_DIR, FONTS_DIR
        import aspose.words as aw
        from pathlib import Path
        import re
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('Hello World!')
        html_fixed_save_options = aw.saving.HtmlFixedSaveOptions()
        html_fixed_save_options.encoding = 'US-ASCII'
        self.assertEqual('US-ASCII', html_fixed_save_options.encoding)
        doc.save(file_name=ARTIFACTS_DIR + 'HtmlFixedSaveOptions.UseEncoding.html', save_options=html_fixed_save_options)
        #ExEnd
        assert re.search('content="text/html; charset=us-ascii"', Path(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.UseEncoding.html').read_text()).group() is not None

    def test_get_encoding(self):
        doc = document_helper.DocumentHelper.create_document_fill_with_dummy_text()
        html_fixed_save_options = aw.saving.HtmlFixedSaveOptions()
        html_fixed_save_options.encoding = system_helper.text.Encoding.utf_8()
        doc.save(file_name=ARTIFACTS_DIR + 'HtmlFixedSaveOptions.GetEncoding.html', save_options=html_fixed_save_options)

    @unittest.skipIf(sys.platform.startswith('win'), 'Discrepancy in assertion between Python and .Net')
    def test_export_embedded_css(self):
        from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, GOLDS_DIR, TEMP_DIR, IMAGE_DIR, FONTS_DIR
        import aspose.words as aw
        import aspose.words.saving as aw_saving
        import system_helper
        import re
        for export_embedded_css in [True, False]:
            #ExStart
            #ExFor:HtmlFixedSaveOptions.export_embedded_css
            #ExSummary:Shows how to determine where to store CSS stylesheets when exporting a document to Html.
            doc = aw.Document(MY_DIR + 'Rendering.docx')
            # When you export a document to html, Aspose.Words will also create a CSS stylesheet to format the document with.
            # Setting the "ExportEmbeddedCss" flag to "true" save the CSS stylesheet to a .css file,
            # and link to the file from the html document using a <link> element.
            # Setting the flag to "false" will embed the CSS stylesheet within the Html document,
            # which will create only one file instead of two.
            html_fixed_save_options = aw_saving.HtmlFixedSaveOptions()
            html_fixed_save_options.export_embedded_css = export_embedded_css
            doc.save(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.ExportEmbeddedCss.html', save_options=html_fixed_save_options)
            out_doc_contents = system_helper.io.File.read_all_text(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.ExportEmbeddedCss.html')
            if export_embedded_css:
                assert re.search('<style type="text/css">', out_doc_contents) is not None
                assert not system_helper.io.File.exist(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.ExportEmbeddedCss/styles.css')
            else:
                assert re.search('<link rel="stylesheet" type="text/css" href="HtmlFixedSaveOptions[.]ExportEmbeddedCss/styles[.]css" media="all" />', out_doc_contents) is not None
                assert system_helper.io.File.exist(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.ExportEmbeddedCss/styles.css')
            #ExEnd

    @unittest.skipIf(sys.platform.startswith('win'), 'Discrepancy in assertion between Python and .Net')
    def test_export_embedded_fonts(self):
        from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, GOLDS_DIR, TEMP_DIR, IMAGE_DIR, FONTS_DIR
        import re
        import aspose.words as aw
        for export_embedded_fonts in [True, False]:
            #ExStart
            #ExFor:HtmlFixedSaveOptions.export_embedded_fonts
            #ExSummary:Shows how to determine where to store embedded fonts when exporting a document to Html.
            doc = aw.Document(MY_DIR + 'Embedded font.docx')
            html_fixed_save_options = aw.saving.HtmlFixedSaveOptions()
            html_fixed_save_options.export_embedded_fonts = export_embedded_fonts
            doc.save(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.ExportEmbeddedFonts.html', save_options=html_fixed_save_options)
            out_doc_contents = system_helper.io.File.read_all_text(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.ExportEmbeddedFonts/styles.css')
            if export_embedded_fonts:
                assert re.search("@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], url[(].+[)] format[(]'woff'[)]; }", out_doc_contents)
                self.assertEqual(0, len(list(filter(lambda f: f.endswith('.woff'), system_helper.io.Directory.get_files(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.ExportEmbeddedFonts')))))
            else:
                assert re.search("@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], url[(]'font001[.]woff'[)] format[(]'woff'[)]; }", out_doc_contents)
                self.assertEqual(2, len(list(filter(lambda f: f.endswith('.woff'), system_helper.io.Directory.get_files(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.ExportEmbeddedFonts')))))
            #ExEnd

    @unittest.skipIf(sys.platform.startswith('win'), 'Discrepancy in assertion between Python and .Net')
    def test_export_embedded_images(self):
        from pathlib import Path
        import re
        from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, GOLDS_DIR, TEMP_DIR, IMAGE_DIR, FONTS_DIR
        import aspose.words as aw
        for export_images in [True, False]:
            #ExStart
            #ExFor:HtmlFixedSaveOptions.export_embedded_images
            #ExSummary:Shows how to determine where to store images when exporting a document to Html.
            doc = aw.Document(MY_DIR + 'Images.docx')
            html_fixed_save_options = aw.saving.HtmlFixedSaveOptions()
            html_fixed_save_options.export_embedded_images = export_images
            doc.save(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.ExportEmbeddedImages.html', save_options=html_fixed_save_options)
            out_doc_contents = system_helper.io.File.read_all_text(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.ExportEmbeddedImages.html')
            if export_images:
                self.assertFalse(system_helper.io.File.exist(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.ExportEmbeddedImages/image001.jpeg'))
                self.assertTrue(re.match('<img class=\\"awimg\\" style=\\"left:0pt; top:0pt; width:493.1pt; height:300.55pt;\\" src=\\".+\\" />', out_doc_contents))
            else:
                self.assertTrue(system_helper.io.File.exist(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.ExportEmbeddedImages/image001.jpeg'))
                self.assertTrue(re.match('<img class=\\"awimg\\" style=\\"left:0pt; top:0pt; width:493.1pt; height:300.55pt;\\" src=\\"HtmlFixedSaveOptions[.]ExportEmbeddedImages/image001[.]jpeg\\" />', out_doc_contents))
    #ExEnd

    def test_export_embedded_svgs(self):
        from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR
        import aspose.words as aw
        import system_helper
        import re

        class ExHtmlFixedSaveOptions(ApiExampleBase):

            def test_export_embedded_svgs(self):
                for export_svgs in [True, False]:
                    #ExStart
                    #ExFor:HtmlFixedSaveOptions.export_embedded_svg
                    #ExSummary:Shows how to determine where to store SVG objects when exporting a document to Html.
                    doc = aw.Document(file_name=MY_DIR + 'Images.docx')
                    # When we export a document with SVG objects to .html,
                    # Aspose.Words can place these objects in two possible locations.
                    # Setting the "ExportEmbeddedSvg" flag to "true" will embed all SVG object raw data
                    # within the output HTML, inside <image> tags.
                    # Setting this flag to "false" will create a file in the local file system for each SVG object.
                    # The HTML will link to each file using the "data" attribute of an <object> tag.
                    html_fixed_save_options = aw.saving.HtmlFixedSaveOptions()
                    html_fixed_save_options.export_embedded_svg = export_svgs
                    doc.save(file_name=ARTIFACTS_DIR + 'HtmlFixedSaveOptions.ExportEmbeddedSvgs.html', save_options=html_fixed_save_options)
                    out_doc_contents = system_helper.io.File.read_all_text(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.ExportEmbeddedSvgs.html')
                    if export_svgs:
                        self.assertFalse(system_helper.io.File.exist(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.ExportEmbeddedSvgs/svg001.svg'))
                        self.assertTrue(re.compile('<image id=\\"image004\\" xlink:href=.+/>').search(out_doc_contents) is not None)
                    else:
                        self.assertTrue(system_helper.io.File.exist(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.ExportEmbeddedSvgs/svg001.svg'))
                        self.assertTrue(re.compile('<object type=\\"image/svg\\+xml\\" data=\\"HtmlFixedSaveOptions\\.ExportEmbeddedSvgs/svg001\\.svg\\"></object>').search(out_doc_contents) is not None)
                    #ExEnd

    def test_add_css_class_names_prefix(self):
        #ExStart
        #ExFor:HtmlFixedSaveOptions.css_class_names_prefix
        #ExFor:HtmlFixedSaveOptions.save_font_face_css_separately
        #ExSummary:Shows how to place CSS into a separate file and add a prefix to all of its CSS class names.
        from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, GOLDS_DIR, TEMP_DIR, IMAGE_DIR, FONTS_DIR
        import aspose.words as aw
        import system_helper
        import re
        from unittest import TestCase
        Assert = TestCase()
        doc = aw.Document(MY_DIR + 'Bookmarks.docx')
        html_fixed_save_options = aw.saving.HtmlFixedSaveOptions()
        html_fixed_save_options.css_class_names_prefix = 'myprefix'
        html_fixed_save_options.save_font_face_css_separately = True
        doc.save(file_name=ARTIFACTS_DIR + 'HtmlFixedSaveOptions.AddCssClassNamesPrefix.html', save_options=html_fixed_save_options)
        out_doc_contents = system_helper.io.File.read_all_text(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.AddCssClassNamesPrefix.html')
        Assert.assertTrue(re.search('<div class="myprefixdiv myprefixpage" style="width:595[.]3pt; height:841[.]9pt;">' + '<div class="myprefixdiv" style="left:85[.]05pt; top:36pt; clip:rect[(]0pt,510[.]25pt,74[.]95pt,-85.05pt[)];">' + '<span class="myprefixspan myprefixtext001" style="font-size:11pt; left:294[.]73pt; top:0[.]36pt; line-height:12[.]29pt;">', out_doc_contents) is not None)
        out_doc_contents = system_helper.io.File.read_all_text(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.AddCssClassNamesPrefix/styles.css')
        Assert.assertTrue(re.search('\\.myprefixdiv \\{ position:absolute; \\} ' + '\\.myprefixspan \\{ position:absolute; white-space:pre; color:#000000; font-size:12pt; \\}', out_doc_contents) is not None)
        #ExEnd
    def test_horizontal_alignment(self):
        import re
        from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR
        import aspose.words as aw
        for page_horizontal_alignment in [aw.saving.HtmlFixedPageHorizontalAlignment.CENTER, aw.saving.HtmlFixedPageHorizontalAlignment.LEFT, aw.saving.HtmlFixedPageHorizontalAlignment.RIGHT]:
            #ExStart
            #ExFor:HtmlFixedSaveOptions.page_horizontal_alignment
            #ExFor:HtmlFixedPageHorizontalAlignment
            #ExSummary:Shows how to set the horizontal alignment of pages when saving a document to HTML.
            doc = aw.Document(file_name=MY_DIR + 'Rendering.docx')
            html_fixed_save_options = aw.saving.HtmlFixedSaveOptions()
            html_fixed_save_options.page_horizontal_alignment = page_horizontal_alignment
            doc.save(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.HorizontalAlignment.html', save_options=html_fixed_save_options)
            out_doc_contents = system_helper.io.File.read_all_text(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.HorizontalAlignment/styles.css')
            switch_condition = page_horizontal_alignment
            if switch_condition == aw.saving.HtmlFixedPageHorizontalAlignment.CENTER:
                assert re.search('\\.awpage \\{ position:relative; border:solid 1pt black; margin:10pt auto 10pt auto; overflow:hidden; \\}', out_doc_contents)
            elif switch_condition == aw.saving.HtmlFixedPageHorizontalAlignment.LEFT:
                assert re.search('\\.awpage \\{ position:relative; border:solid 1pt black; margin:10pt auto 10pt 10pt; overflow:hidden; \\}', out_doc_contents)
            elif switch_condition == aw.saving.HtmlFixedPageHorizontalAlignment.RIGHT:
                assert re.search('\\.awpage \\{ position:relative; border:solid 1pt black; margin:10pt 10pt 10pt auto; overflow:hidden; \\}', out_doc_contents)
            #ExEnd

    @unittest.skipIf(sys.platform.startswith('win'), 'Discrepancy in assertion between Python and .Net')
    def test_page_margins(self):
        import re
        from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, GOLDS_DIR, TEMP_DIR, IMAGE_DIR, FONTS_DIR
        import aspose.words as aw
        from pathlib import Path
        doc = aw.Document(MY_DIR + 'Document.docx')
        save_options = aw.saving.HtmlFixedSaveOptions()
        save_options.page_margins = 15
        doc.save(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.PageMargins.html', save_options=save_options)
        out_doc_contents = system_helper.io.File.read_all_text(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.PageMargins/styles.css')
        assert re.match('[.]awpage { position:relative; border:solid 1pt black; margin:15pt auto 15pt auto; overflow:hidden; }', out_doc_contents) is not None

    def test_page_margins_exception(self):
        save_options = aw.saving.HtmlFixedSaveOptions()
        with self.assertRaises(Exception):
            save_options.page_margins = -1

    def test_optimize_graphics_output(self):
        from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR
        from pathlib import Path
        for optimize_output in [False, True]:
            #ExStart
            #ExFor:HtmlFixedSaveOptions.optimize_output
            #ExSummary:Shows how to simplify a document when saving it to HTML by removing various redundant objects.
            doc = aw.Document(file_name=MY_DIR + 'Rendering.docx')
            save_options = aw.saving.HtmlFixedSaveOptions()
            save_options.optimize_output = optimize_output
            doc.save(file_name=ARTIFACTS_DIR + 'HtmlFixedSaveOptions.OptimizeGraphicsOutput.html', save_options=save_options)
            # The size of the optimized version of the document is almost a third of the size of the unoptimized document.
            self.assertAlmostEqual(60385 if optimize_output else 191000, Path(ARTIFACTS_DIR + 'HtmlFixedSaveOptions.OptimizeGraphicsOutput.html').stat().st_size, delta=200)
            #ExEnd

    def _test_resource_saving_callback(self, callback):
        self.assertTrue('font001.woff' in callback.get_text())
        self.assertTrue('styles.css' in callback.get_text())

    def _test_html_fixed_resource_folder(self, callback):
        from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, GOLDS_DIR, TEMP_DIR, IMAGE_DIR, FONTS_DIR
        import re
        from aspose import words as aw
        # Replace Assert.AreEqual with standard Python assertion
        assert re.findall('Resource #', callback.get_text()).__len__() == 16
        # Fix for the NameError: Assert is not defined in Python
        # Also convert C# Regex.Matches().Count to Python's re.findall().__len__()
        assert re.match('[.]awpage { position:relative; border:solid 1pt black; margin:15pt auto 15pt auto; overflow:hidden; }', out_doc_contents) is not None

    def test_id_prefix(self):
        #ExStart:IdPrefix
        #ExFor:HtmlFixedSaveOptions.id_prefix
        #ExSummary:Shows how to add a prefix that is prepended to all generated element IDs.
        doc = aw.Document(file_name=MY_DIR + 'Id prefix.docx')
        save_options = aw.saving.HtmlFixedSaveOptions()
        save_options.id_prefix = 'pfx1_'
        doc.save(file_name=ARTIFACTS_DIR + 'HtmlFixedSaveOptions.IdPrefix.html', save_options=save_options)
        #ExEnd:IdPrefix

    def test_remove_java_script_from_links(self):
        #ExStart:RemoveJavaScriptFromLinks
        #ExFor:HtmlFixedSaveOptions.remove_java_script_from_links
        #ExSummary:Shows how to remove JavaScript from the links for html fixed documents.
        doc = aw.Document(file_name=MY_DIR + 'JavaScript in HREF.docx')
        save_options = aw.saving.HtmlFixedSaveOptions()
        save_options.remove_java_script_from_links = True
        doc.save(file_name=ARTIFACTS_DIR + 'HtmlFixedSaveOptions.RemoveJavaScriptFromLinks.html', save_options=save_options)
        #ExEnd:RemoveJavaScriptFromLinks
    #ExStart
    #ExFor:IResourceSavingCallback
    #ExFor:IResourceSavingCallback.resource_saving(ResourceSavingArgs)
    #ExFor:ResourceSavingArgs
    #ExFor:ResourceSavingArgs.document
    #ExFor:ResourceSavingArgs.resource_file_name
    #ExFor:ResourceSavingArgs.resource_file_uri
    #ExSummary:Shows how to use a callback to track external resources created while converting a document to HTML (FontSavingCallback).
    class FontSavingCallback(aw.saving.IResourceSavingCallback):

        def __init__(self):
            self.m_text = []

        def resource_saving(self, args):
            self.m_text.append(f'Original document URI:\t{args.document.original_file_name}' + '\n')
            self.m_text.append(f'Resource being saved:\t{args.resource_file_name}' + '\n')
            self.m_text.append(f'Full uri after saving:\t{args.resource_file_uri}\n' + '\n')

        def get_text(self):
            return str.join('', self.m_text)
    #ExEnd
    #ExStart
    #ExFor:HtmlFixedSaveOptions
    #ExFor:HtmlFixedSaveOptions.resource_saving_callback
    #ExFor:HtmlFixedSaveOptions.resources_folder
    #ExFor:HtmlFixedSaveOptions.resources_folder_alias
    #ExFor:HtmlFixedSaveOptions.save_format
    #ExFor:HtmlFixedSaveOptions.show_page_border
    #ExFor:IResourceSavingCallback
    #ExFor:IResourceSavingCallback.resource_saving(ResourceSavingArgs)
    #ExFor:ResourceSavingArgs.keep_resource_stream_open
    #ExFor:ResourceSavingArgs.resource_stream
    #ExSummary:Shows how to use a callback to print the URIs of external resources created while converting a document to HTML (ResourceUriPrinter).

    class ResourceUriPrinter(aw.saving.IResourceSavingCallback):

        def __init__(self):
            self.m_saved_resource_count = None
            self.m_text = []

        def resource_saving(self, args):
            # If we set a folder alias in the SaveOptions object, we will be able to print it from here.
            self.m_saved_resource_count += 1
            self.m_text.append(f'Resource #{self.m_saved_resource_count} "{args.resource_file_name}"')
            extension = Path(args.resource_file_name).suffix
            if extension in ('.ttf', '.woff'):
                # By default, 'ResourceFileUri' uses system folder for fonts.
                # To avoid problems in other platforms you must explicitly specify the path for the fonts.
                args.resource_file_uri = str(Path(ARTIFACTS_DIR) / args.resource_file_name)
            self.m_text.append('\t' + args.resource_file_uri + '\n')
            # If we have specified a folder in the "ResourcesFolderAlias" property,
            # we will also need to redirect each stream to put its resource in that folder.
            args.resource_stream = system_helper.io.FileStream(args.resource_file_uri, system_helper.io.FileMode.CREATE)
            args.keep_resource_stream_open = False

        def get_text(self):
            return str.join('', self.m_text)
    #ExEnd