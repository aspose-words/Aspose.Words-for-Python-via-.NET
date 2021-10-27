import unittest
import os
import sys

from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

import aspose.words as aw

class WorkingWithHtmlSaveOptions(DocsExamplesBase):

    def test_export_roundtrip_information(self):

        #ExStart:ExportRoundtripInformation
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.HtmlSaveOptions()
        save_options.export_roundtrip_information = True

        doc.save(ARTIFACTS_DIR + "WorkingWithHtmlSaveOptions.export_roundtrip_information.html", save_options)
        #ExEnd:ExportRoundtripInformation

    def test_export_fonts_as_base_64(self):

        #ExStart:ExportFontsAsBase64
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.HtmlSaveOptions()
        save_options.export_fonts_as_base64 = True

        doc.save(ARTIFACTS_DIR + "WorkingWithHtmlSaveOptions.export_fonts_as_base64.html", save_options)
        #ExEnd:ExportFontsAsBase64

    def test_export_resources(self):

        #ExStart:ExportResources
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.HtmlSaveOptions()

        save_options.css_style_sheet_type = aw.saving.CssStyleSheetType.EXTERNAL
        save_options.export_font_resources = True
        save_options.resource_folder = ARTIFACTS_DIR + "Resources"
        save_options.resource_folder_alias = "http://example.com/resources"

        doc.save(ARTIFACTS_DIR + "WorkingWithHtmlSaveOptions.export_resources.html", save_options)
        #ExEnd:ExportResources

    def test_convert_metafiles_to_emf_or_wmf(self):

        #ExStart:ConvertMetafilesToEmfOrWmf
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Here is an image as is: ")
        builder.insert_html("""<img src="data:image/pngbase64,
            iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAABGdBTUEAALGP
            C/xhBQAAAAlwSFlzAAALEwAACxMBAJqcGAAAAAd0SU1FB9YGARc5KB0XV+IA
            AAAddEVYdENvbW1lbnQAQ3JlYXRlZCB3aXRoIFRoZSBHSU1Q72QlbgAAAF1J
            REFUGNO9zL0NglAAxPEfdLTs4BZM4DIO4C7OwQg2JoQ9LE1exdlYvBBeZ7jq
            ch9#q1uH4TLzw4d6+ErXMMcXuHWxId3KOETnnXXV6MJpcq2MLaI97CER3N0
            vr4MkhoXe0rZigAAAABJRU5ErkJggg==" alt="Red dot" />""")

        save_options = aw.saving.HtmlSaveOptions()
        save_options.metafile_format = aw.saving.HtmlMetafileFormat.EMF_OR_WMF

        doc.save(ARTIFACTS_DIR + "WorkingWithHtmlSaveOptions.convert_metafiles_to_emf_or_wmf.html", save_options)
        #ExEnd:ConvertMetafilesToEmfOrWmf

    def test_convert_metafiles_to_svg(self):

        #ExStart:ConvertMetafilesToSvg
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Here is an SVG image: ")
        builder.insert_html("""
            <svg height='210' width='500'>
                <polygon points='100,10 40,198 190,78 10,78 160,198'
                    style='fill:limestroke:purplestroke-width:5fill-rule:evenodd' />
            </svg>""")

        save_options = aw.saving.HtmlSaveOptions()
        save_options.metafile_format = aw.saving.HtmlMetafileFormat.SVG

        doc.save(ARTIFACTS_DIR + "WorkingWithHtmlSaveOptions.convert_metafiles_to_svg.html", save_options)
        #ExEnd:ConvertMetafilesToSvg

    def test_add_css_class_name_prefix(self):

        #ExStart:AddCssClassNamePrefix
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.HtmlSaveOptions()
        save_options.css_style_sheet_type = aw.saving.CssStyleSheetType.EXTERNAL
        save_options.css_class_name_prefix = "pfx_"

        doc.save(ARTIFACTS_DIR + "WorkingWithHtmlSaveOptions.add_css_class_name_prefix.html", save_options)
        #ExEnd:AddCssClassNamePrefix

    def test_export_cid_urls_for_mhtml_resources(self):

        #ExStart:ExportCidUrlsForMhtmlResources
        doc = aw.Document(MY_DIR + "Content-ID.docx")

        save_options = aw.saving.HtmlSaveOptions(aw.SaveFormat.MHTML)
        save_options.pretty_format = True
        save_options.export_cid_urls_for_mhtml_resources = True

        doc.save(ARTIFACTS_DIR + "WorkingWithHtmlSaveOptions.export_cid_urls_for_mhtml_resources.mhtml", save_options)
        #ExEnd:ExportCidUrlsForMhtmlResources

    def test_resolve_font_names(self):

        #ExStart:ResolveFontNames
        doc = aw.Document(MY_DIR + "Missing font.docx")

        save_options = aw.saving.HtmlSaveOptions(aw.SaveFormat.HTML)
        save_options.pretty_format = True
        save_options.resolve_font_names = True

        doc.save(ARTIFACTS_DIR + "WorkingWithHtmlSaveOptions.resolve_font_names.html", save_options)
        #ExEnd:ResolveFontNames

    def test_export_text_input_form_field_as_text(self):

        #ExStart:ExportTextInputFormFieldAsText
        doc = aw.Document(MY_DIR + "Rendering.docx")

        images_dir = os.path.join(ARTIFACTS_DIR, "Images")

        # The folder specified needs to exist and should be empty.
        if os.path.exists(images_dir):
            os.rmdir(images_dir)

        os.makedirs(images_dir)

        # Set an option to export form fields as plain text, not as HTML input elements.
        save_options = aw.saving.HtmlSaveOptions(aw.SaveFormat.HTML)
        save_options.export_text_input_form_field_as_text = True
        save_options.images_folder = images_dir

        doc.save(ARTIFACTS_DIR + "WorkingWithHtmlSaveOptions.export_text_input_form_field_as_text.html", save_options)
        #ExEnd:ExportTextInputFormFieldAsText

    def test_convert_document_to_epub(self):

        #ExStart:ConvertDocumentToEPUB
        # Load the document from disk.
        doc = aw.Document(MY_DIR + "Rendering.docx")

        # Create a new instance of HtmlSaveOptions. This object allows us to set options that control
        # How the output document is saved.
        save_options = aw.saving.HtmlSaveOptions()

        # Specify the desired encoding.
        save_options.encoding = "utf-8"

        # Specify at what elements to split the internal HTML at. This creates a new HTML within the EPUB
        # which allows you to limit the size of each HTML part. This is useful for readers which cannot read
        # HTML files greater than a certain size e.g 300kb.
        save_options.document_split_criteria = aw.saving.DocumentSplitCriteria.HEADING_PARAGRAPH

        # Specify that we want to export document properties.
        save_options.export_document_properties = True

        # Specify that we want to save in EPUB format.
        save_options.save_format = aw.SaveFormat.EPUB

        # Export the document as an EPUB file.
        doc.save(ARTIFACTS_DIR + "Document.EpubConversion_out.epub", save_options)
        #ExEnd:ConvertDocumentToEPUB
