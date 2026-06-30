import io

import aspose.words as aw
import aspose.pydrawing as drawing
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class WorkingWithImageSaveOptions(DocsExamplesBase):

    def test_expose_threshold_control(self):

        #ExStart:ExposeThresholdControl
        #GistId:b20a0ec0e1ff0556aa20d12f486e1963
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.TIFF)

        save_options.tiff_compression = aw.saving.TiffCompression.CCITT3
        save_options.image_color_mode = aw.saving.ImageColorMode.GRAYSCALE
        save_options.tiff_binarization_method = aw.saving.ImageBinarizationMethod.FLOYD_STEINBERG_DITHERING
        save_options.threshold_for_floyd_steinberg_dithering = 254

        doc.save(ARTIFACTS_DIR + "WorkingWithImageSaveOptions.expose_threshold_control.tiff", save_options)
        #ExEnd:ExposeThresholdControl

    def test_get_tiff_page_range(self):

        #ExStart:GetTiffPageRange
        #GistId:b20a0ec0e1ff0556aa20d12f486e1963
        doc = aw.Document(MY_DIR + "Rendering.docx")
        #ExStart:SaveAsTiff
        #GistId:b20a0ec0e1ff0556aa20d12f486e1963
        doc.save(ARTIFACTS_DIR + "WorkingWithImageSaveOptions.multipage_tiff.tiff")
        #ExEnd:SaveAsTiff

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
        #GistId:83e5c469d0e72b5114fb8a05a1d01977
        doc = aw.Document(MY_DIR + "Rendering.docx")

        save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)

        save_options.page_set = aw.saving.PageSet(1)
        save_options.image_color_mode = aw.saving.ImageColorMode.BLACK_AND_WHITE
        save_options.pixel_format = aw.saving.ImagePixelFormat.FORMAT_1BPP_INDEXED

        doc.save(ARTIFACTS_DIR + "WorkingWithImageSaveOptions.format_1_bpp_indexed.png", save_options)
        #ExEnd:Format1BppIndexed

    def test_get_jpeg_page_range(self):

        #ExStart:GetJpegPageRange
        #GistId:ebbb90d74ef57db456685052a18f8e86
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

    def test_horizontal_layout(self):

        #ExStart:HorizontalLayout
        #GistId:8eeaafcfcc55d78505f0f378ad8c6907
        doc = aw.Document(MY_DIR + "Rendering.docx")

        options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)
        options.page_layout = aw.saving.MultiPageLayout.horizontal(10)

        doc.save(ARTIFACTS_DIR + "WorkingWithImageSaveOptions.horizontal_layout.jpg", options)
        #ExEnd:HorizontalLayout

    def test_grid_layout(self):

        #ExStart:GridLayout
        #GistId:8eeaafcfcc55d78505f0f378ad8c6907
        doc = aw.Document(MY_DIR + "Rendering.docx")

        options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)
        # Set up a grid layout with:
        # - 3 columns per row.
        # - 10pts spacing between pages (horizontal and vertical).
        options.page_layout = aw.saving.MultiPageLayout.grid(3, 10, 10)

        # Customize the background and border.
        options.page_layout.back_color = drawing.Color.light_gray
        options.page_layout.border_color = drawing.Color.blue
        options.page_layout.border_width = 2

        doc.save(ARTIFACTS_DIR + "WorkingWithImageSaveOptions.grid_layout.jpg", options)
        #ExEnd:GridLayout
