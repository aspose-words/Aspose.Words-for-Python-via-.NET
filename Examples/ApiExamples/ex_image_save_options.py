# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import sys
import glob
import os
import aspose.pydrawing
import aspose.words as aw
import aspose.words.saving
import system_helper
import test_util
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, IMAGE_DIR, MY_DIR

class ExImageSaveOptions(ApiExampleBase):

    def test_one_page(self):
        #ExStart
        #ExFor:Document.save(str,SaveOptions)
        #ExFor:FixedPageSaveOptions
        #ExFor:ImageSaveOptions.page_set
        #ExFor:PageSet
        #ExFor:PageSet.__init__(int)
        #ExSummary:Shows how to render one page from a document to a JPEG image.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('Page 1.')
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln('Page 2.')
        builder.insert_image(file_name=IMAGE_DIR + 'Logo.jpg')
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln('Page 3.')
        # Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
        # to modify the way in which that method renders the document into an image.
        options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)
        # Set the "PageSet" to "1" to select the second page via
        # the zero-based index to start rendering the document from.
        options.page_set = aw.saving.PageSet(page=1)
        # When we save the document to the JPEG format, Aspose.Words only renders one page.
        # This image will contain one page starting from page two,
        # which will just be the second page of the original document.
        doc.save(file_name=ARTIFACTS_DIR + 'ImageSaveOptions.OnePage.jpg', save_options=options)
        #ExEnd
        test_util.TestUtil.verify_image(816, 1056, ARTIFACTS_DIR + 'ImageSaveOptions.OnePage.jpg')

    @unittest.skipUnless(sys.platform.startswith('win'), 'different calculation on Linux')
    def test_renderer(self):
        for use_gdi_emf_renderer in [False, True]:
            #ExStart
            #ExFor:ImageSaveOptions.use_gdi_emf_renderer
            #ExSummary:Shows how to choose a renderer when converting a document to .emf.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.paragraph_format.style = doc.styles.get_by_name('Heading 1')
            builder.writeln('Hello world!')
            builder.insert_image(file_name=IMAGE_DIR + 'Logo.jpg')
            # When we save the document as an EMF image, we can pass a SaveOptions object to select a renderer for the image.
            # If we set the "UseGdiEmfRenderer" flag to "true", Aspose.Words will use the GDI+ renderer.
            # If we set the "UseGdiEmfRenderer" flag to "false", Aspose.Words will use its own metafile renderer.
            save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.EMF)
            save_options.use_gdi_emf_renderer = use_gdi_emf_renderer
            doc.save(file_name=ARTIFACTS_DIR + 'ImageSaveOptions.Renderer.emf', save_options=save_options)
            #ExEnd

    def test_page_set(self):
        #ExStart
        #ExFor:ImageSaveOptions.page_set
        #ExSummary:Shows how to specify which page in a document to render as an image.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.paragraph_format.style = doc.styles.get_by_name('Heading 1')
        builder.writeln('Hello world! This is page 1.')
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln('This is page 2.')
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln('This is page 3.')
        self.assertEqual(3, doc.page_count)
        # When we save the document as an image, Aspose.Words only renders the first page by default.
        # We can pass a SaveOptions object to specify a different page to render.
        save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.GIF)
        # Render every page of the document to a separate image file.
        i = 1
        while i <= doc.page_count:
            save_options.page_set = aw.saving.PageSet(page=1)
            doc.save(file_name=ARTIFACTS_DIR + f'ImageSaveOptions.PageIndex.Page {i}.gif', save_options=save_options)
            i += 1
        #ExEnd
        test_util.TestUtil.verify_image(816, 1056, ARTIFACTS_DIR + 'ImageSaveOptions.PageIndex.Page 1.gif')
        test_util.TestUtil.verify_image(816, 1056, ARTIFACTS_DIR + 'ImageSaveOptions.PageIndex.Page 2.gif')
        test_util.TestUtil.verify_image(816, 1056, ARTIFACTS_DIR + 'ImageSaveOptions.PageIndex.Page 3.gif')
        self.assertFalse(system_helper.io.File.exist(ARTIFACTS_DIR + 'ImageSaveOptions.PageIndex.Page 4.gif'))

    def test_windows_meta_file(self):
        for metafile_rendering_mode in [aw.saving.MetafileRenderingMode.VECTOR, aw.saving.MetafileRenderingMode.BITMAP, aw.saving.MetafileRenderingMode.VECTOR_WITH_FALLBACK]:
            #ExStart
            #ExFor:ImageSaveOptions.metafile_rendering_options
            #ExFor:MetafileRenderingOptions.use_gdi_raster_operations_emulation
            #ExSummary:Shows how to set the rendering mode when saving documents with Windows Metafile images to other image formats.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.insert_image(file_name=IMAGE_DIR + 'Windows MetaFile.wmf')
            # When we save the document as an image, we can pass a SaveOptions object to
            # determine how the saving operation will process Windows Metafiles in the document.
            # If we set the "RenderingMode" property to "MetafileRenderingMode.Vector",
            # or "MetafileRenderingMode.VectorWithFallback", we will render all metafiles as vector graphics.
            # If we set the "RenderingMode" property to "MetafileRenderingMode.Bitmap", we will render all metafiles as bitmaps.
            options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
            options.metafile_rendering_options.rendering_mode = metafile_rendering_mode
            # Aspose.Words uses GDI+ for raster operations emulation, when value is set to true.
            options.metafile_rendering_options.use_gdi_raster_operations_emulation = True
            doc.save(file_name=ARTIFACTS_DIR + 'ImageSaveOptions.WindowsMetaFile.png', save_options=options)
            #ExEnd
            test_util.TestUtil.verify_image(816, 1056, ARTIFACTS_DIR + 'ImageSaveOptions.WindowsMetaFile.png')

    @unittest.skipUnless(sys.platform.startswith('win'), 'different calculation on Mac')
    def test_color_mode(self):
        for image_color_mode in [aw.saving.ImageColorMode.BLACK_AND_WHITE, aw.saving.ImageColorMode.GRAYSCALE, aw.saving.ImageColorMode.NONE]:
            #ExStart
            #ExFor:ImageColorMode
            #ExFor:ImageSaveOptions.image_color_mode
            #ExSummary:Shows how to set a color mode when rendering documents.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.paragraph_format.style = doc.styles.get_by_name('Heading 1')
            builder.writeln('Hello world!')
            builder.insert_image(file_name=IMAGE_DIR + 'Logo.jpg')
            # When we save the document as an image, we can pass a SaveOptions object to
            # select a color mode for the image that the saving operation will generate.
            # If we set the "ImageColorMode" property to "ImageColorMode.BlackAndWhite",
            # the saving operation will apply grayscale color reduction while rendering the document.
            # If we set the "ImageColorMode" property to "ImageColorMode.Grayscale",
            # the saving operation will render the document into a monochrome image.
            # If we set the "ImageColorMode" property to "None", the saving operation will apply the default method
            # and preserve all the document's colors in the output image.
            image_save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
            image_save_options.image_color_mode = image_color_mode
            doc.save(file_name=ARTIFACTS_DIR + 'ImageSaveOptions.ColorMode.png', save_options=image_save_options)
            #ExEnd
            tested_image_length = system_helper.io.FileInfo(ARTIFACTS_DIR + 'ImageSaveOptions.ColorMode.png').length()
            switch_condition = image_color_mode
            if switch_condition == aw.saving.ImageColorMode.NONE:
                self.assertTrue(tested_image_length < 175000)
            elif switch_condition == aw.saving.ImageColorMode.GRAYSCALE:
                self.assertTrue(tested_image_length < 90000)
            elif switch_condition == aw.saving.ImageColorMode.BLACK_AND_WHITE:
                self.assertTrue(tested_image_length < 15000)

    def test_floyd_steinberg_dithering(self):
        #ExStart
        #ExFor:ImageBinarizationMethod
        #ExFor:ImageSaveOptions.threshold_for_floyd_steinberg_dithering
        #ExFor:ImageSaveOptions.tiff_binarization_method
        #ExSummary:Shows how to set the TIFF binarization error threshold when using the Floyd-Steinberg method to render a TIFF image.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.paragraph_format.style = doc.styles.get_by_name('Heading 1')
        builder.writeln('Hello world!')
        builder.insert_image(file_name=IMAGE_DIR + 'Logo.jpg')
        # When we save the document as a TIFF, we can pass a SaveOptions object to
        # adjust the dithering that Aspose.Words will apply when rendering this image.
        # The default value of the "ThresholdForFloydSteinbergDithering" property is 128.
        # Higher values tend to produce darker images.
        options = aw.saving.ImageSaveOptions(aw.SaveFormat.TIFF)
        options.tiff_compression = aw.saving.TiffCompression.CCITT3
        options.tiff_binarization_method = aw.saving.ImageBinarizationMethod.FLOYD_STEINBERG_DITHERING
        options.threshold_for_floyd_steinberg_dithering = 240
        doc.save(file_name=ARTIFACTS_DIR + 'ImageSaveOptions.FloydSteinbergDithering.tiff', save_options=options)
        #ExEnd
        image_file_names = list(filter(lambda item: 'ImageSaveOptions.FloydSteinbergDithering.' in item and item.endswith('.tiff'), list(system_helper.io.Directory.get_files(ARTIFACTS_DIR, '*.tiff'))))
        self.assertEqual(1, len(image_file_names))

    @unittest.skip("Discrepancy in assertion between Python and .Net")
    def test_edit_image(self):
        #ExStart
        #ExFor:ImageSaveOptions.horizontal_resolution
        #ExFor:ImageSaveOptions.image_brightness
        #ExFor:ImageSaveOptions.image_contrast
        #ExFor:ImageSaveOptions.save_format
        #ExFor:ImageSaveOptions.scale
        #ExFor:ImageSaveOptions.vertical_resolution
        #ExSummary:Shows how to edit the image while Aspose.Words converts a document to one.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.paragraph_format.style = doc.styles.get_by_name('Heading 1')
        builder.writeln('Hello world!')
        builder.insert_image(file_name=IMAGE_DIR + 'Logo.jpg')
        # When we save the document as an image, we can pass a SaveOptions object to
        # edit the image while the saving operation renders it.
        options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
        # We can adjust these properties to change the image's brightness and contrast.
        # Both are on a 0-1 scale and are at 0.5 by default.
        options.image_brightness = 0.3
        options.image_contrast = 0.7
        # We can adjust horizontal and vertical resolution with these properties.
        # This will affect the dimensions of the image.
        # The default value for these properties is 96.0, for a resolution of 96dpi.
        options.horizontal_resolution = 72
        options.vertical_resolution = 72
        # We can scale the image using this property. The default value is 1.0, for scaling of 100%.
        # We can use this property to negate any changes in image dimensions that changing the resolution would cause.
        options.scale = 96 / 72
        doc.save(file_name=ARTIFACTS_DIR + 'ImageSaveOptions.EditImage.png', save_options=options)
        #ExEnd
        test_util.TestUtil.verify_image(817, 1057, ARTIFACTS_DIR + 'ImageSaveOptions.EditImage.png')

    def test_jpeg_quality(self):
        #ExStart
        #ExFor:Document.save(str,SaveOptions)
        #ExFor:FixedPageSaveOptions.jpeg_quality
        #ExFor:ImageSaveOptions
        #ExFor:ImageSaveOptions.__init__
        #ExFor:ImageSaveOptions.jpeg_quality
        #ExSummary:Shows how to configure compression while saving a document as a JPEG.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.insert_image(file_name=IMAGE_DIR + 'Logo.jpg')
        # Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
        # to modify the way in which that method renders the document into an image.
        image_options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)
        # Set the "JpegQuality" property to "10" to use stronger compression when rendering the document.
        # This will reduce the file size of the document, but the image will display more prominent compression artifacts.
        image_options.jpeg_quality = 10
        doc.save(file_name=ARTIFACTS_DIR + 'ImageSaveOptions.JpegQuality.HighCompression.jpg', save_options=image_options)
        # Set the "JpegQuality" property to "100" to use weaker compression when rending the document.
        # This will improve the quality of the image at the cost of an increased file size.
        image_options.jpeg_quality = 100
        doc.save(file_name=ARTIFACTS_DIR + 'ImageSaveOptions.JpegQuality.HighQuality.jpg', save_options=image_options)
        #ExEnd
        self.assertTrue(system_helper.io.FileInfo(ARTIFACTS_DIR + 'ImageSaveOptions.JpegQuality.HighCompression.jpg').length() < 18000)
        self.assertTrue(system_helper.io.FileInfo(ARTIFACTS_DIR + 'ImageSaveOptions.JpegQuality.HighQuality.jpg').length() < 75000)

    def test_tiff_image_compression(self):
        for tiff_compression in [aw.saving.TiffCompression.NONE, aw.saving.TiffCompression.RLE, aw.saving.TiffCompression.LZW, aw.saving.TiffCompression.CCITT3, aw.saving.TiffCompression.CCITT4]:
            #ExStart
            #ExFor:TiffCompression
            #ExFor:ImageSaveOptions.tiff_compression
            #ExSummary:Shows how to select the compression scheme to apply to a document that we convert into a TIFF image.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.insert_image(file_name=IMAGE_DIR + 'Logo.jpg')
            # Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
            # to modify the way in which that method renders the document into an image.
            options = aw.saving.ImageSaveOptions(aw.SaveFormat.TIFF)
            # Set the "TiffCompression" property to "TiffCompression.None" to apply no compression while saving,
            # which may result in a very large output file.
            # Set the "TiffCompression" property to "TiffCompression.Rle" to apply RLE compression
            # Set the "TiffCompression" property to "TiffCompression.Lzw" to apply LZW compression.
            # Set the "TiffCompression" property to "TiffCompression.Ccitt3" to apply CCITT3 compression.
            # Set the "TiffCompression" property to "TiffCompression.Ccitt4" to apply CCITT4 compression.
            options.tiff_compression = tiff_compression
            doc.save(file_name=ARTIFACTS_DIR + 'ImageSaveOptions.TiffImageCompression.tiff', save_options=options)
            #ExEnd
            tested_image_length = system_helper.io.FileInfo(ARTIFACTS_DIR + 'ImageSaveOptions.TiffImageCompression.tiff').length()
            switch_condition = tiff_compression
            if switch_condition == aw.saving.TiffCompression.NONE:
                self.assertTrue(tested_image_length < 3450000)
            elif switch_condition == aw.saving.TiffCompression.RLE:
                self.assertTrue(tested_image_length < 687000)
            elif switch_condition == aw.saving.TiffCompression.LZW:
                self.assertTrue(tested_image_length < 250000)
            elif switch_condition == aw.saving.TiffCompression.CCITT3:
                self.assertTrue(tested_image_length < 8300)
            elif switch_condition == aw.saving.TiffCompression.CCITT4:
                self.assertTrue(tested_image_length < 1700)

    def test_export_various_page_ranges(self):
        #ExStart
        #ExFor:PageSet.__init__(List[PageRange])
        #ExFor:PageRange
        #ExFor:PageRange.__init__(int,int)
        #ExFor:ImageSaveOptions.page_set
        #ExSummary:Shows how to extract pages based on exact page ranges.
        doc = aw.Document(file_name=MY_DIR + 'Images.docx')
        image_options = aw.saving.ImageSaveOptions(aw.SaveFormat.TIFF)
        page_set = aw.saving.PageSet(ranges=[aw.saving.PageRange(1, 1), aw.saving.PageRange(2, 3), aw.saving.PageRange(1, 3), aw.saving.PageRange(2, 4), aw.saving.PageRange(1, 1)])
        image_options.page_set = page_set
        doc.save(file_name=ARTIFACTS_DIR + 'ImageSaveOptions.ExportVariousPageRanges.tiff', save_options=image_options)
        #ExEnd

    def test_render_ink_object(self):
        #ExStart
        #ExFor:SaveOptions.iml_rendering_mode
        #ExFor:ImlRenderingMode
        #ExSummary:Shows how to render Ink object.
        doc = aw.Document(file_name=MY_DIR + 'Ink object.docx')
        # Set 'ImlRenderingMode.InkML' ignores fall-back shape of ink (InkML) object and renders InkML itself.
        # If the rendering result is unsatisfactory,
        # please use 'ImlRenderingMode.Fallback' to get a result similar to previous versions.
        save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)
        save_options.iml_rendering_mode = aw.saving.ImlRenderingMode.INK_ML
        doc.save(file_name=ARTIFACTS_DIR + 'ImageSaveOptions.RenderInkObject.jpeg', save_options=save_options)
        #ExEnd

    def test_grid_layout(self):
        #ExStart:GridLayout
        #ExFor:ImageSaveOptions.page_layout
        #ExFor:MultiPageLayout
        #ExSummary:Shows how to save the document into JPG image with multi-page layout settings.
        doc = aw.Document(file_name=MY_DIR + 'Rendering.docx')
        options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)
        # Set up a grid layout with:
        # - 3 columns per row.
        # - 10pts spacing between pages (horizontal and vertical).
        options.page_layout = aw.saving.MultiPageLayout.grid(3, 10, 10)
        # Alternative layouts:
        # options.PageLayout = MultiPageLayout.Horizontal(10);
        # options.PageLayout = MultiPageLayout.Vertical(10);
        # Customize the background and border.
        options.page_layout.back_color = aspose.pydrawing.Color.light_gray
        options.page_layout.border_color = aspose.pydrawing.Color.blue
        options.page_layout.border_width = 2
        doc.save(file_name=ARTIFACTS_DIR + 'ImageSaveOptions.GridLayout.jpg', save_options=options)
        #ExEnd:GridLayout

    def test_page_by_page(self):
        #ExStart
        #ExFor:Document.save(str,SaveOptions)
        #ExFor:FixedPageSaveOptions
        #ExFor:ImageSaveOptions.page_set
        #ExFor: ImageSaveOptions.image_size
        #ExSummary:Shows how to render every page of a document to a separate TIFF image.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln('Page 1.')
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln('Page 2.')
        builder.insert_image(IMAGE_DIR + 'Logo.jpg')
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln('Page 3.')
        # Create an "ImageSaveOptions" object which we can pass to the document's "save" method
        # to modify the way in which that method renders the document into an image.
        options = aw.saving.ImageSaveOptions(aw.SaveFormat.TIFF)
        for i in range(doc.page_count):
            # Set the "page_set" property to the number of the first page from
            # which to start rendering the document from.
            options.page_set = aw.saving.PageSet(i)
            options.vertical_resolution = 600
            options.horizontal_resolution = 600
            options.image_size = aspose.pydrawing.Size(2325, 5325)
            doc.save(ARTIFACTS_DIR + f'ImageSaveOptions.page_by_page.{i + 1}.tiff', options)
        #ExEnd
        image_file_names = glob.glob(ARTIFACTS_DIR + '/ImageSaveOptions.page_by_page*.tiff')
        self.assertEqual(3, len(image_file_names))
        for image_file_name in image_file_names:
            self.verify_image(2325, 5325, filename=image_file_name)

    @unittest.skip("Discrepancy in assertion between Python and .Net")
    def test_pixel_format(self):
        for image_pixel_format in [aw.saving.ImagePixelFormat.FORMAT_1BPP_INDEXED, aw.saving.ImagePixelFormat.FORMAT_16BPP_RGB_555, aw.saving.ImagePixelFormat.FORMAT_16BPP_RGB_565, aw.saving.ImagePixelFormat.FORMAT_24BPP_RGB, aw.saving.ImagePixelFormat.FORMAT_32BPP_RGB, aw.saving.ImagePixelFormat.FORMAT_32BPP_ARGB, aw.saving.ImagePixelFormat.FORMAT_32BPP_P_ARGB, aw.saving.ImagePixelFormat.FORMAT_48BPP_RGB, aw.saving.ImagePixelFormat.FORMAT_64BPP_ARGB, aw.saving.ImagePixelFormat.FORMAT_64BPP_P_ARGB]:
            #ExStart
            #ExFor:ImagePixelFormat
            #ExFor:ImageSaveOptions.clone
            #ExFor:ImageSaveOptions.pixel_format
            #ExSummary:Shows how to select a bit-per-pixel rate with which to render a document to an image.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.paragraph_format.style = doc.styles.get_by_name('Heading 1')
            builder.writeln('Hello world!')
            builder.insert_image(file_name=IMAGE_DIR + 'Logo.jpg')
            # When we save the document as an image, we can pass a SaveOptions object to
            # select a pixel format for the image that the saving operation will generate.
            # Various bit per pixel rates will affect the quality and file size of the generated image.
            image_save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
            image_save_options.pixel_format = image_pixel_format
            # We can clone ImageSaveOptions instances.
            self.assertNotEqual(image_save_options, image_save_options.clone())
            doc.save(file_name=ARTIFACTS_DIR + 'ImageSaveOptions.PixelFormat.png', save_options=image_save_options)
            #ExEnd
            tested_image_length = system_helper.io.FileInfo(ARTIFACTS_DIR + 'ImageSaveOptions.PixelFormat.png').length()
            switch_condition = image_pixel_format
            if switch_condition == aw.saving.ImagePixelFormat.FORMAT_1BPP_INDEXED:
                self.assertTrue(tested_image_length < 10000)
            elif switch_condition == aw.saving.ImagePixelFormat.FORMAT_16BPP_RGB_565:
                self.assertTrue(tested_image_length < 150000)
            elif switch_condition == aw.saving.ImagePixelFormat.FORMAT_16BPP_RGB_555:
                self.assertTrue(tested_image_length < 150000)
            elif switch_condition == aw.saving.ImagePixelFormat.FORMAT_24BPP_RGB:
                self.assertTrue(tested_image_length < 90000)
            elif switch_condition == aw.saving.ImagePixelFormat.FORMAT_32BPP_RGB or switch_condition == aw.saving.ImagePixelFormat.FORMAT_32BPP_ARGB:
                self.assertTrue(tested_image_length < 150000)
            elif switch_condition == aw.saving.ImagePixelFormat.FORMAT_48BPP_RGB:
                self.assertTrue(tested_image_length < 150000)
            elif switch_condition == aw.saving.ImagePixelFormat.FORMAT_64BPP_ARGB or switch_condition == aw.saving.ImagePixelFormat.FORMAT_64BPP_P_ARGB:
                self.assertTrue(tested_image_length < 150000)

    def test_resolution(self):
        #ExStart
        #ExFor:ImageSaveOptions
        #ExFor:ImageSaveOptions.resolution
        #ExSummary:Shows how to specify a resolution while rendering a document to PNG.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.name = 'Times New Roman'
        builder.font.size = 24
        builder.writeln('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.')
        builder.insert_image(file_name=IMAGE_DIR + 'Logo.jpg')
        # Create an "ImageSaveOptions" object which we can pass to the document's "Save" method
        # to modify the way in which that method renders the document into an image.
        options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
        # Set the "Resolution" property to "72" to render the document in 72dpi.
        options.vertical_resolution = 72
        options.horizontal_resolution = 72
        doc.save(file_name=ARTIFACTS_DIR + 'ImageSaveOptions.Resolution.72dpi.png', save_options=options)
        # Set the "Resolution" property to "300" to render the document in 300dpi.
        options.vertical_resolution = 300
        options.horizontal_resolution = 300
        doc.save(file_name=ARTIFACTS_DIR + 'ImageSaveOptions.Resolution.300dpi.png', save_options=options)
        #ExEnd
        test_util.TestUtil.verify_image(612, 792, ARTIFACTS_DIR + 'ImageSaveOptions.Resolution.72dpi.png')
        test_util.TestUtil.verify_image(2550, 3300, ARTIFACTS_DIR + 'ImageSaveOptions.Resolution.300dpi.png')