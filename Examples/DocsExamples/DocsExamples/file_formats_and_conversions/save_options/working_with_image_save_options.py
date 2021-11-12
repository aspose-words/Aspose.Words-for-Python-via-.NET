import io

import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class WorkingWithImageSaveOptions(DocsExamplesBase):

    def test_expose_threshold_control_for_tiff_binarization(self):

        #ExStart:ExposeThresholdControlForTiffBinarization
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.TIFF)

        save_options.tiff_compression = aw.saving.TiffCompression.CCITT3
        save_options.image_color_mode = aw.saving.ImageColorMode.GRAYSCALE
        save_options.tiff_binarization_method = aw.saving.ImageBinarizationMethod.FLOYD_STEINBERG_DITHERING
        save_options.threshold_for_floyd_steinberg_dithering = 254

        doc.save(ARTIFACTS_DIR + "WorkingWithImageSaveOptions.expose_threshold_control_for_tiff_binarization.tiff", save_options)
        #ExEnd:ExposeThresholdControlForTiffBinarization

    def test_get_tiff_page_range(self):

        #ExStart:GetTiffPageRange
        doc = aw.Document(MY_DIR + "Rendering.docx")
        #ExStart:SaveAsTIFF
        doc.save(ARTIFACTS_DIR + "WorkingWithImageSaveOptions.multipage_tiff.tiff")
        #ExEnd:SaveAsTIFF

        #ExStart:SaveAsTIFFUsingImageSaveOptions
        save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.TIFF)
        save_options.page_set = aw.saving.PageSet([0, 1])
        save_options.tiff_compression = aw.saving.TiffCompression.CCITT4
        save_options.vertical_resolution = 160
        save_options.horizontal_resolution = 160

        doc.save(ARTIFACTS_DIR + "WorkingWithImageSaveOptions.get_tiff_page_range.tiff", save_options)
        #ExEnd:SaveAsTIFFUsingImageSaveOptions
        #ExEnd:GetTiffPageRange

    def test_format_1_bpp_indexed(self):

        #ExStart:Format1BppIndexed
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)

        save_options.page_set = aw.saving.PageSet(1)
        save_options.image_color_mode = aw.saving.ImageColorMode.BLACK_AND_WHITE
        save_options.pixel_format = aw.saving.ImagePixelFormat.FORMAT1BPP_INDEXED

        doc.save(ARTIFACTS_DIR + "WorkingWithImageSaveOptions.format_1_bpp_indexed.png", save_options)
        #ExEnd:Format1BppIndexed

    def test_get_jpeg_page_range(self):

        #ExStart:GetJpegPageRange
        doc = aw.Document(MY_DIR + "Rendering.docx")

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

        doc.save(ARTIFACTS_DIR + "WorkingWithImageSaveOptions.get_jpeg_page_range.jpeg", options)
        #ExEnd:GetJpegPageRange

    def test_save_document_to_jpeg(self):

        #ExStart:SaveDocumentToJPEG
        # Open the document
        doc = aw.Document(MY_DIR + "Rendering.docx")
        # Save as a JPEG image file with default options
        doc.save(ARTIFACTS_DIR + "Rendering.JpegDefaultOptions.jpg")

        # Save document to stream as a JPEG with default options
        doc_stream = io.BytesIO()
        doc.save(doc_stream, aw.SaveFormat.JPEG)
        # Rewind the stream position back to the beginning, ready for use
        doc_stream.seek(0)

        # Save document to a JPEG image with specified options.
        # Render the third page only and set the JPEG quality to 80%
        # In this case we need to pass the desired SaveFormat to the ImageSaveOptions constructor
        # to signal what type of image to save as.
        image_options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)
        image_options.page_set = aw.saving.PageSet(2)
        image_options.jpeg_quality = 80
        doc.save(ARTIFACTS_DIR + "Rendering.JpegCustomOptions.jpg", image_options)
        #ExEnd:SaveDocumentToJPEG
