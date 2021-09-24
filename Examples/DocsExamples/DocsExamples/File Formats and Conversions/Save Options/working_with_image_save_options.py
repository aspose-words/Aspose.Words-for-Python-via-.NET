import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithImageSaveOptions(docs_base.DocsExamplesBase):
    
    def test_expose_threshold_control_for_tiff_binarization(self) :
        
        #ExStart:ExposeThresholdControlForTiffBinarization
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        saveOptions = aw.saving.ImageSaveOptions(aw.SaveFormat.TIFF)
            
        saveOptions.tiff_compression = aw.saving.TiffCompression.CCITT3
        saveOptions.image_color_mode = aw.saving.ImageColorMode.GRAYSCALE
        saveOptions.tiff_binarization_method = aw.saving.ImageBinarizationMethod.FLOYD_STEINBERG_DITHERING
        saveOptions.threshold_for_floyd_steinberg_dithering = 254
            

        doc.save(docs_base.artifacts_dir + "WorkingWithImageSaveOptions.expose_threshold_control_for_tiff_binarization.tiff", saveOptions)
        #ExEnd:ExposeThresholdControlForTiffBinarization
        

    def test_get_tiff_page_range(self) :
        
        #ExStart:GetTiffPageRange
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")
        #ExStart:SaveAsTIFF
        doc.save(docs_base.artifacts_dir + "WorkingWithImageSaveOptions.multipage_tiff.tiff")
        #ExEnd:SaveAsTIFF
            
        #ExStart:SaveAsTIFFUsingImageSaveOptions
        saveOptions = aw.saving.ImageSaveOptions(aw.SaveFormat.TIFF)
        saveOptions.page_set = aw.saving.PageSet([0, 1])
        saveOptions.tiff_compression = aw.saving.TiffCompression.CCITT4
        #saveOptions.resolution = 160
            

        doc.save(docs_base.artifacts_dir + "WorkingWithImageSaveOptions.get_tiff_page_range.tiff", saveOptions)
        #ExEnd:SaveAsTIFFUsingImageSaveOptions
        #ExEnd:GetTiffPageRange
        

    def test_format_1_bpp_indexed(self) :
        
        #ExStart:Format1BppIndexed
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        saveOptions = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
            
        saveOptions.page_set = aw.saving.PageSet(1)
        saveOptions.image_color_mode = aw.saving.ImageColorMode.BLACK_AND_WHITE
        saveOptions.pixel_format = aw.saving.ImagePixelFormat.FORMAT1BPP_INDEXED
            

        doc.save(docs_base.artifacts_dir + "WorkingWithImageSaveOptions.format_1_bpp_indexed.png", saveOptions)
        #ExEnd:Format1BppIndexed
        

    def test_get_jpeg_page_range(self) :
        
        #ExStart:GetJpegPageRange
        doc = aw.Document(docs_base.my_dir + "Rendering.docx")

        options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)

        # Set the "PageSet" to "0" to convert only the first page of a document.
        options.page_set = aw.saving.PageSet(0)

        # Change the image's brightness and contrast.
        # Both are on a 0-1 scale and are at 0.5 by default.
        options.image_brightness = 0.3
        options.image_contrast = 0.7

        # Change the horizontal resolution.
        # The default value for these properties is 96.0, for a resolution of 96dpi.
        options.horizontal_resolution = 72

        doc.save(docs_base.artifacts_dir + "WorkingWithImageSaveOptions.get_jpeg_page_range.jpeg", options)
        #ExEnd:GetJpegPageRange


if __name__ == '__main__':
    unittest.main()