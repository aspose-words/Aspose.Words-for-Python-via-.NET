import unittest
import os
import sys
from datetime import date, datetime

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithPdfSaveOptions(docs_base.DocsExamplesBase):
    
    def test_display_doc_title_in_window_titlebar(self) :
        
        #ExStart:DisplayDocTitleInWindowTitlebar
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.display_doc_title = True 

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.display_doc_title_in_window_titlebar.pdf", saveOptions)
        #ExEnd:DisplayDocTitleInWindowTitlebar
        

    def test_digitally_signed_pdf_using_certificate_holder(self) :
        
        #ExStart:DigitallySignedPdfUsingCertificateHolder
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
            
        builder.writeln("Test Signed PDF.")

        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.digital_signature_details = aw.saving.PdfDigitalSignatureDetails(
                aw.digitalsignatures.CertificateHolder.create(docs_base.my_dir + "morzal.pfx", "aw"), "reason", "location", datetime.today())
            

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.digitally_signed_pdf_using_certificate_holder.pdf", saveOptions)
        #ExEnd:DigitallySignedPdfUsingCertificateHolder
        

    def test_embedded_all_fonts(self) :
        
        #ExStart:EmbeddAllFonts
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        # The output PDF will be embedded with all fonts found in the document.
        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.embed_full_fonts = True 
            
        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.embedded_fonts_in_pdf.pdf", saveOptions)
        #ExEnd:EmbeddAllFonts
        

    def test_embedded_subset_fonts(self) :
        
        #ExStart:EmbeddSubsetFonts
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        # The output PDF will contain subsets of the fonts in the document.
        # Only the glyphs used in the document are included in the PDF fonts.
        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.embed_full_fonts = False 
            
        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.embedd_subset_fonts.pdf", saveOptions)
        #ExEnd:EmbeddSubsetFonts
        

    def test_disable_embed_windows_fonts(self) :
        
        #ExStart:DisableEmbedWindowsFonts
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        # The output PDF will be saved without embedding standard windows fonts.
        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.font_embedding_mode = aw.saving.PdfFontEmbeddingMode.EMBED_NONE 
            
        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.disable_embed_windows_fonts.pdf", saveOptions)
        #ExEnd:DisableEmbedWindowsFonts
        

    def test_skip_embedded_arial_and_times_roman_fonts(self) :
        
        #ExStart:SkipEmbeddedArialAndTimesRomanFonts
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.font_embedding_mode =  aw.saving.PdfFontEmbeddingMode.EMBED_ALL 

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.skip_embedded_arial_and_times_roman_fonts.pdf", saveOptions)
        #ExEnd:SkipEmbeddedArialAndTimesRomanFonts
        

    def test_avoid_embedding_core_fonts(self) :
        
        #ExStart:AvoidEmbeddingCoreFonts
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        # The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.use_core_fonts = True 
            
        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.avoid_embedding_core_fonts.pdf", saveOptions)
        #ExEnd:AvoidEmbeddingCoreFonts
        
        
    def test_escape_uri(self) :
        
        #ExStart:EscapeUri
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
            
        builder.insert_hyperlink("Testlink", "https:#www.google.com/search?q=%2Fthe%20test", False)
        builder.writeln()
        builder.insert_hyperlink("https:#www.google.com/search?q=%2Fthe%20test", "https:#www.google.com/search?q=%2Fthe%20test", False)

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.escape_uri.pdf")
        #ExEnd:EscapeUri
        

    def test_export_header_footer_bookmarks(self) :
        
        #ExStart:ExportHeaderFooterBookmarks
        doc = aw.Document(docs_base.my_dir + "Bookmarks in headers and footers.docx")

        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.outline_options.default_bookmarks_outline_level = 1
        saveOptions.header_footer_bookmarks_export_mode = aw.saving.HeaderFooterBookmarksExportMode.FIRST

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.export_header_footer_bookmarks.pdf", saveOptions)
        #ExEnd:ExportHeaderFooterBookmarks
        

    def test_scale_wmf_fonts_to_metafile_size(self) :
        
        #ExStart:ScaleWmfFontsToMetafileSize
        doc = aw.Document(docs_base.my_dir + "WMF with text.docx")

        metafileRenderingOptions = aw.saving.MetafileRenderingOptions()
        metafileRenderingOptions.scale_wmf_fonts_to_metafile_size = False
            

        # If Aspose.words cannot correctly render some of the metafile records to vector graphics
        # then Aspose.words renders this metafile to a bitmap.
        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.metafile_rendering_options = metafileRenderingOptions 

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.scale_wmf_fonts_to_metafile_size.pdf", saveOptions)
        #ExEnd:ScaleWmfFontsToMetafileSize
        

    def test_additional_text_positioning(self) :
        
        #ExStart:AdditionalTextPositioning
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.additional_text_positioning = True 

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.additional_text_positioning.pdf", saveOptions)
        #ExEnd:AdditionalTextPositioning
        

    def test_conversion_to_pdf_17(self) :
        
        #ExStart:ConversionToPDF17
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.compliance = aw.saving.PdfCompliance.PDF17 

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.conversion_to_pdf_17.pdf", saveOptions)
        #ExEnd:ConversionToPDF17
        

    def test_downsampling_images(self) :
        
        #ExStart:DownsamplingImages
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        # We can set a minimum threshold for downsampling.
        # This value will prevent the second image in the input document from being downsampled.
        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.downsample_options.resolution = 36
        saveOptions.downsample_options.resolution_threshold = 128 

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.downsampling_images.pdf", saveOptions)
        #ExEnd:DownsamplingImages
        

    def test_set_outline_options(self) :
        
        #ExStart:SetOutlineOptions
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.outline_options.headings_outline_levels = 3
        saveOptions.outline_options.expanded_outline_levels = 1

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.set_outline_options.pdf", saveOptions)
        #ExEnd:SetOutlineOptions
        

    def test_custom_properties_export(self) :
        
        #ExStart:CustomPropertiesExport
        doc = aw.Document()
        doc.custom_document_properties.add("Company", "Aspose")

        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.custom_properties_export = aw.saving.PdfCustomPropertiesExport.STANDARD 

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.custom_properties_export.pdf", saveOptions)
        #ExEnd:CustomPropertiesExport
        

    def test_export_document_structure(self) :
        
        #ExStart:ExportDocumentStructure
        doc = aw.Document(docs_base.my_dir + "Paragraphs.docx")

        # The file size will be increased and the structure will be visible in the "Content" navigation pane
        # of Adobe Acrobat Pro, while editing the .pdf.
        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.export_document_structure = True 

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.export_document_structure.pdf", saveOptions)
        #ExEnd:ExportDocumentStructure
        

    def test_image_compression(self) :
        
        #ExStart:PdfImageComppression
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.image_compression = aw.saving.PdfImageCompression.JPEG
        saveOptions.preserve_form_fields = True

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.pdf_image_compression.pdf", saveOptions)

        saveOptionsA2U = aw.saving.PdfSaveOptions
        saveOptionsA2U.compliance = aw.saving.PdfCompliance.PDFA2U
        saveOptionsA2U.image_compression = aw.saving.PdfImageCompression.JPEG
        saveOptionsA2U.jpeg_quality = 100 # Use JPEG compression at 50% quality to reduce file size.

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.pdf_image_compression__a_2u.pdf", saveOptionsA2U)
        #ExEnd:PdfImageComppression
        

    def test_update_last_printed_property(self) :
        
        #ExStart:UpdateIfLastPrinted
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.update_last_printed_property = True 

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.update_if_last_printed.pdf", saveOptions)
        #ExEnd:UpdateIfLastPrinted
        

    def test_dml_3_d_effects_rendering(self) :
        
        #ExStart:Dml3DEffectsRendering
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.dml3_deffects_rendering_mode = aw.saving.Dml3DEffectsRenderingMode.ADVANCED 

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.dml_3_d_effects_rendering.pdf", saveOptions)
        #ExEnd:Dml3DEffectsRendering
        

    def test_interpolate_images(self) :
        
        #ExStart:SetImageInterpolation
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.interpolate_images = True 

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfSaveOptions.interpolate_images.pdf", saveOptions)
        #ExEnd:SetImageInterpolation
        
    

if __name__ == '__main__':
    unittest.main()