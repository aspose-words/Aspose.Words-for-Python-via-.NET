from datetime import datetime
import re

import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class WorkingWithPdfSaveOptions(DocsExamplesBase):

    def test_display_doc_title_in_window_titlebar(self):

        #ExStart:DisplayDocTitleInWindowTitlebar
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.PdfSaveOptions()
        save_options.display_doc_title = True

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.display_doc_title_in_window_titlebar.pdf", save_options)
        #ExEnd:DisplayDocTitleInWindowTitlebar

    def test_digitally_signed_pdf_using_certificate_holder(self):

        #ExStart:DigitallySignedPdfUsingCertificateHolder
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Test Signed PDF.")

        certificate = aw.digitalsignatures.CertificateHolder.create(MY_DIR + "morzal.pfx", "aw")

        save_options = aw.saving.PdfSaveOptions()
        save_options.digital_signature_details = aw.saving.PdfDigitalSignatureDetails(
                certificate, "reason", "location", datetime.now())

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.digitally_signed_pdf_using_certificate_holder.pdf", save_options)
        #ExEnd:DigitallySignedPdfUsingCertificateHolder

    def test_embedded_all_fonts(self):

        #ExStart:EmbeddAllFonts
        doc = aw.Document(MY_DIR + "Rendering.docx")

        # The output PDF will be embedded with all fonts found in the document.
        save_options = aw.saving.PdfSaveOptions()
        save_options.embed_full_fonts = True

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.embedded_fonts_in_pdf.pdf", save_options)
        #ExEnd:EmbeddAllFonts

    def test_embedded_subset_fonts(self):

        #ExStart:EmbeddSubsetFonts
        doc = aw.Document(MY_DIR + "Rendering.docx")

        # The output PDF will contain subsets of the fonts in the document.
        # Only the glyphs used in the document are included in the PDF fonts.
        save_options = aw.saving.PdfSaveOptions()
        save_options.embed_full_fonts = False

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.embedd_subset_fonts.pdf", save_options)
        #ExEnd:EmbeddSubsetFonts

    def test_disable_embed_windows_fonts(self):

        #ExStart:DisableEmbedWindowsFonts
        doc = aw.Document(MY_DIR + "Rendering.docx")

        # The output PDF will be saved without embedding standard windows fonts.
        save_options = aw.saving.PdfSaveOptions()
        save_options.font_embedding_mode = aw.saving.PdfFontEmbeddingMode.EMBED_NONE

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.disable_embed_windows_fonts.pdf", save_options)
        #ExEnd:DisableEmbedWindowsFonts

    def test_skip_embedded_arial_and_times_roman_fonts(self):

        #ExStart:SkipEmbeddedArialAndTimesRomanFonts
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.PdfSaveOptions()
        save_options.font_embedding_mode =  aw.saving.PdfFontEmbeddingMode.EMBED_ALL

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.skip_embedded_arial_and_times_roman_fonts.pdf", save_options)
        #ExEnd:SkipEmbeddedArialAndTimesRomanFonts

    def test_avoid_embedding_core_fonts(self):

        #ExStart:AvoidEmbeddingCoreFonts
        doc = aw.Document(MY_DIR + "Rendering.docx")

        # The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
        save_options = aw.saving.PdfSaveOptions()
        save_options.use_core_fonts = True

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.avoid_embedding_core_fonts.pdf", save_options)
        #ExEnd:AvoidEmbeddingCoreFonts

    def test_escape_uri(self):

        #ExStart:EscapeUri
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_hyperlink("Testlink",
                                 "https://www.google.com/search?q=%2Fthe%20test", False)
        builder.writeln()
        builder.insert_hyperlink("https://www.google.com/search?q=%2Fthe%20test",
                                 "https://www.google.com/search?q=%2Fthe%20test", False)

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.escape_uri.pdf")
        #ExEnd:EscapeUri

    def test_export_header_footer_bookmarks(self):

        #ExStart:ExportHeaderFooterBookmarks
        doc = aw.Document(MY_DIR + "Bookmarks in headers and footers.docx")

        save_options = aw.saving.PdfSaveOptions()
        save_options.outline_options.default_bookmarks_outline_level = 1
        save_options.header_footer_bookmarks_export_mode = aw.saving.HeaderFooterBookmarksExportMode.FIRST

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.export_header_footer_bookmarks.pdf", save_options)
        #ExEnd:ExportHeaderFooterBookmarks

    def test_emulate_rendering_to_size_on_page(self):


        #ExStart:EmulateRenderingToSizeOnPage
        doc = aw.Document(MY_DIR + "WMF with text.docx")

        metafile_rendering_options = aw.saving.MetafileRenderingOptions()
        metafile_rendering_options.emulate_rendering_to_size_on_page = False

        # If Aspose.Words cannot correctly render some of the metafile records to vector graphics
        # then Aspose.Words renders this metafile to a bitmap.
        save_options = aw.saving.PdfSaveOptions()
        save_options.metafile_rendering_options = metafile_rendering_options

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.EmulateRenderingToSizeOnPage.pdf", save_options)
        #ExEnd:EmulateRenderingToSizeOnPage

    def test_additional_text_positioning(self):

        #ExStart:AdditionalTextPositioning
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.PdfSaveOptions()
        save_options.additional_text_positioning = True

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.additional_text_positioning.pdf", save_options)
        #ExEnd:AdditionalTextPositioning

    def test_conversion_to_pdf_17(self):

        #ExStart:ConversionToPdf17
        #GistId:36a49a29062268dc5e6d3134163f8d99
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.PdfSaveOptions()
        save_options.compliance = aw.saving.PdfCompliance.PDF17

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.conversion_to_pdf_17.pdf", save_options)
        #ExEnd:ConversionToPdf17

    def test_downsampling_images(self):

        #ExStart:DownsamplingImages
        doc = aw.Document(MY_DIR + "Rendering.docx")

        # We can set a minimum threshold for downsampling.
        # This value will prevent the second image in the input document from being downsampled.
        save_options = aw.saving.PdfSaveOptions()
        save_options.downsample_options.resolution = 36
        save_options.downsample_options.resolution_threshold = 128

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.downsampling_images.pdf", save_options)
        #ExEnd:DownsamplingImages

    def test_set_outline_options(self):

        #ExStart:SetOutlineOptions
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.PdfSaveOptions()
        save_options.outline_options.headings_outline_levels = 3
        save_options.outline_options.expanded_outline_levels = 1

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.set_outline_options.pdf", save_options)
        #ExEnd:SetOutlineOptions

    def test_custom_properties_export(self):

        #ExStart:CustomPropertiesExport
        doc = aw.Document()
        doc.custom_document_properties.add("Company", "Aspose")

        save_options = aw.saving.PdfSaveOptions()
        save_options.custom_properties_export = aw.saving.PdfCustomPropertiesExport.STANDARD

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.custom_properties_export.pdf", save_options)
        #ExEnd:CustomPropertiesExport

    def test_export_document_structure(self):

        #ExStart:ExportDocumentStructure
        doc = aw.Document(MY_DIR + "Paragraphs.docx")

        # The file size will be increased and the structure will be visible in the "Content" navigation pane
        # of Adobe Acrobat Pro, while editing the .pdf.
        save_options = aw.saving.PdfSaveOptions()
        save_options.export_document_structure = True

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.export_document_structure.pdf", save_options)
        #ExEnd:ExportDocumentStructure

    def test_image_compression(self):

        #ExStart:PdfImageComppression
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.PdfSaveOptions()
        save_options.image_compression = aw.saving.PdfImageCompression.JPEG
        save_options.preserve_form_fields = True

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.pdf_image_compression.pdf", save_options)

        save_options_a2u = aw.saving.PdfSaveOptions()
        save_options_a2u.compliance = aw.saving.PdfCompliance.PDF_A2U
        save_options_a2u.image_compression = aw.saving.PdfImageCompression.JPEG
        save_options_a2u.jpeg_quality = 100 # Use JPEG compression at 50% quality to reduce file size.

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.pdf_image_compression_a2u.pdf", save_options_a2u)
        #ExEnd:PdfImageComppression

    def test_update_last_printed_property(self):

        #ExStart:UpdateIfLastPrinted
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.PdfSaveOptions()
        save_options.update_last_printed_property = True

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.update_if_last_printed.pdf", save_options)
        #ExEnd:UpdateIfLastPrinted

    def test_dml3d_effects_rendering(self):

        #ExStart:Dml3DEffectsRendering
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.PdfSaveOptions()
        save_options.dml_3d_effects_rendering_mode = aw.saving.Dml3DEffectsRenderingMode.ADVANCED

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.dml_3d_effects_rendering.pdf", save_options)
        #ExEnd:Dml3DEffectsRendering

    def test_interpolate_images(self):

        #ExStart:SetImageInterpolation
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.PdfSaveOptions()
        save_options.interpolate_images = True

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfSaveOptions.interpolate_images.pdf", save_options)
        #ExEnd:SetImageInterpolation

    def test_render_metafile_to_bitmap(self):

        #ExStart:RenderMetafileToBitmap
        # Load the document from disk.
        doc = aw.Document(MY_DIR + "Rendering.docx")

        metafile_rendering_options = aw.saving.MetafileRenderingOptions()
        metafile_rendering_options.emulate_raster_operations = False
        metafile_rendering_options.rendering_mode = aw.saving.MetafileRenderingMode.VECTOR_WITH_FALLBACK

        save_options = aw.saving.PdfSaveOptions()
        save_options.metafile_rendering_options = metafile_rendering_options

        doc.save(ARTIFACTS_DIR + "PdfSaveOptions.HandleRasterWarnings.pdf", save_options)
        #ExEnd:RenderMetafileToBitmap

    def test_optimize_output(self):

        #ExStart:OptimizeOutput
        #GistId:36a49a29062268dc5e6d3134163f8d99
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.PdfSaveOptions()
        save_options.optimize_output = True

        doc.save(ARTIFACTS_DIR + "PdfSaveOptions.OptimizeOutput.pdf", save_options)
        #ExEnd:OptimizeOutput
