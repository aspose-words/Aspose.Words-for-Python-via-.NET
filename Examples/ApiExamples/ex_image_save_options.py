# Copyright (c) 2001-2022 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import os
import glob

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, IMAGE_DIR

class ExImageSaveOptions(ApiExampleBase):

    def test_one_page(self):

        #ExStart
        #ExFor:Document.save(str,SaveOptions)
        #ExFor:FixedPageSaveOptions
        #ExFor:ImageSaveOptions.page_set
        #ExSummary:Shows how to render one page from a document to a JPEG image.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Page 1.")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln("Page 2.")
        builder.insert_image(IMAGE_DIR + "Logo.jpg")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln("Page 3.")

        # Create an "ImageSaveOptions" object which we can pass to the document's "save" method
        # to modify the way in which that method renders the document into an image.
        options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)

        # Set the "page_set" to "1" to select the second page via
        # the zero-based index to start rendering the document from.
        options.page_set = aw.saving.PageSet(1)

        # When we save the document to the JPEG format, Aspose.Words only renders one page.
        # This image will contain one page starting from page two,
        # which will just be the second page of the original document.
        doc.save(ARTIFACTS_DIR + "ImageSaveOptions.one_page.jpg", options)
        #ExEnd

        self.verify_image(816, 1056, filename=ARTIFACTS_DIR + "ImageSaveOptions.one_page.jpg")

    def test_renderer(self):

        for use_gdi_emf_renderer in (False, True):
            with self.subTest(use_gdi_emf_renderer=use_gdi_emf_renderer):
                #ExStart
                #ExFor:ImageSaveOptions.use_gdi_emf_renderer
                #ExSummary:Shows how to choose a renderer when converting a document to .emf.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.paragraph_format.style = doc.styles.get_by_name("Heading 1")
                builder.writeln("Hello world!")
                builder.insert_image(IMAGE_DIR + "Logo.jpg")

                # When we save the document as an EMF image, we can pass a SaveOptions object to select a renderer for the image.
                # If we set the "use_gdi_emf_renderer" flag to "True", Aspose.Words will use the GDI+ renderer.
                # If we set the "use_gdi_emf_renderer" flag to "False", Aspose.Words will use its own metafile renderer.
                save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.EMF)
                save_options.use_gdi_emf_renderer = use_gdi_emf_renderer

                doc.save(ARTIFACTS_DIR + "ImageSaveOptions.renderer.emf", save_options)

                # The GDI+ renderer usually creates larger files.
                if use_gdi_emf_renderer:
                    self.assertGreater(30000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.renderer.emf"))
                else:
                    self.assertGreater(30000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.renderer.emf"))
                #ExEnd

                self.verify_image(816, 1056, filename=ARTIFACTS_DIR + "ImageSaveOptions.renderer.emf")

    def test_page_set(self):

        #ExStart
        #ExFor:ImageSaveOptions.page_set
        #ExSummary:Shows how to specify which page in a document to render as an image.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.paragraph_format.style = doc.styles.get_by_name("Heading 1")
        builder.writeln("Hello world! This is page 1.")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln("This is page 2.")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln("This is page 3.")

        self.assertEqual(3, doc.page_count)

        # When we save the document as an image, Aspose.Words only renders the first page by default.
        # We can pass a SaveOptions object to specify a different page to render.
        save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.GIF)

        # Render every page of the document to a separate image file.
        for i in range(1, doc.page_count + 1):
            save_options.page_set = aw.saving.PageSet(1)

            doc.save(ARTIFACTS_DIR + f"ImageSaveOptions.page_set.page {i}.gif", save_options)

        #ExEnd

        self.verify_image(816, 1056, filename=ARTIFACTS_DIR + "ImageSaveOptions.page_set.page 1.gif")
        self.verify_image(816, 1056, filename=ARTIFACTS_DIR + "ImageSaveOptions.page_set.page 2.gif")
        self.verify_image(816, 1056, filename=ARTIFACTS_DIR + "ImageSaveOptions.page_set.page 3.gif")
        self.assertFalse(os.path.exists(ARTIFACTS_DIR + "ImageSaveOptions.page_set.page 4.gif"))

    def test_graphics_quality(self):

        #ExStart
        #ExFor:GraphicsQualityOptions
        #ExFor:GraphicsQualityOptions.compositing_mode
        #ExFor:GraphicsQualityOptions.compositing_quality
        #ExFor:GraphicsQualityOptions.interpolation_mode
        #ExFor:GraphicsQualityOptions.string_format
        #ExFor:GraphicsQualityOptions.smoothing_mode
        #ExFor:GraphicsQualityOptions.text_rendering_hint
        #ExFor:ImageSaveOptions.graphics_quality_options
        #ExSummary:Shows how to set render quality options while converting documents to image formats.
        doc = aw.Document(MY_DIR + "Rendering.docx")

        quality_options = aw.saving.GraphicsQualityOptions()
        quality_options.smoothing_mode = drawing.drawing2d.SmoothingMode.ANTI_ALIAS
        quality_options.text_rendering_hint = drawing.text.text_rendering_hint.CLEAR_TYPE_GRID_FIT
        quality_options.compositing_mode = drawing.drawing2d.CompositingMode.SOURCE_OVER
        quality_options.compositing_quality = drawing.drawing2d.CompositingQuality.HIGH_QUALITY
        quality_options.interpolation_mode = drawing.drawing2d.InterpolationMode.HIGH
        quality_options.string_format = drawing.StringFormat.GENERIC_TYPOGRAPHIC

        save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)
        save_options.graphics_quality_options = quality_options

        doc.save(ARTIFACTS_DIR + "ImageSaveOptions.graphics_quality.jpg", save_options)
        #ExEnd

        self.verify_image(794, 1122, filename=(ARTIFACTS_DIR + "ImageSaveOptions.graphics_quality.jpg"))

    def test_windows_meta_file(self):

        for metafile_rendering_mode in (aw.saving.MetafileRenderingMode.VECTOR,
                                        aw.saving.MetafileRenderingMode.BITMAP,
                                        aw.saving.MetafileRenderingMode.VECTOR_WITH_FALLBACK):
            with self.subTest(metafile_rendering_mode=metafile_rendering_mode):
                #ExStart
                #ExFor:ImageSaveOptions.metafile_rendering_options
                #ExSummary:Shows how to set the rendering mode when saving documents with Windows Metafile images to other image formats.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.insert_image(drawing.Image.from_file(IMAGE_DIR + "Windows MetaFile.wmf"))

                # When we save the document as an image, we can pass a SaveOptions object to
                # determine how the saving operation will process Windows Metafiles in the document.
                # If we set the "rendering_mode" property to "MetafileRenderingMode.VECTOR",
                # or "MetafileRenderingMode.VECTOR_WITH_FALLBACK", we will render all metafiles as vector graphics.
                # If we set the "rendering_mode" property to "MetafileRenderingMode.BITMAP", we will render all metafiles as bitmaps.
                options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
                options.metafile_rendering_options.rendering_mode = metafile_rendering_mode

                doc.save(ARTIFACTS_DIR + "ImageSaveOptions.windows_meta_file.png", options)
                #ExEnd

                self.verify_image(816, 1056, filename=ARTIFACTS_DIR + "ImageSaveOptions.windows_meta_file.png")

    def test_page_by_page(self):

        #ExStart
        #ExFor:Document.save(str,SaveOptions)
        #ExFor:FixedPageSaveOptions
        #ExFor:ImageSaveOptions.page_set
        #ExSummary:Shows how to render every page of a document to a separate TIFF image.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Page 1.")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln("Page 2.")
        builder.insert_image(IMAGE_DIR + "Logo.jpg")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln("Page 3.")

        # Create an "ImageSaveOptions" object which we can pass to the document's "save" method
        # to modify the way in which that method renders the document into an image.
        options = aw.saving.ImageSaveOptions(aw.SaveFormat.TIFF)

        for i in range(doc.page_count):

            # Set the "page_set" property to the number of the first page from
            # which to start rendering the document from.
            options.page_set = aw.saving.PageSet(i)

            doc.save(ARTIFACTS_DIR + f"ImageSaveOptions.page_by_page.{i + 1}.tiff", options)

        #ExEnd

        image_file_names = glob.glob(ARTIFACTS_DIR + "/ImageSaveOptions.page_by_page*.tiff")

        self.assertEqual(3, len(image_file_names))

        for image_file_name in image_file_names:
            self.verify_image(816, 1056, filename=image_file_name)

    def test_color_mode(self):

        for image_color_mode in (aw.saving.ImageColorMode.BLACK_AND_WHITE,
                                 aw.saving.ImageColorMode.GRAYSCALE,
                                 aw.saving.ImageColorMode.NONE):
            with self.subTest(image_color_mode=image_color_mode):
                #ExStart
                #ExFor:ImageColorMode
                #ExFor:ImageSaveOptions.image_color_mode
                #ExSummary:Shows how to set a color mode when rendering documents.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.paragraph_format.style = doc.styles.get_by_name("Heading 1")
                builder.writeln("Hello world!")
                builder.insert_image(IMAGE_DIR + "Logo.jpg")

                self.assertLess(20000, os.path.getsize(IMAGE_DIR + "Logo.jpg"))

                # When we save the document as an image, we can pass a SaveOptions object to
                # select a color mode for the image that the saving operation will generate.
                # If we set the "image_color_mode" property to "ImageColorMode.BLACK_AND_WHITE",
                # the saving operation will apply grayscale color reduction while rendering the document.
                # If we set the "image_color_mode" property to "ImageColorMode.GRAYSCALE",
                # the saving operation will render the document into a monochrome image.
                # If we set the "image_color_mode" property to "NONE", the saving operation will apply the default method
                # and preserve all the document's colors in the output image.
                image_save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
                image_save_options.image_color_mode = image_color_mode

                doc.save(ARTIFACTS_DIR + "ImageSaveOptions.color_mode.png", image_save_options)

                if image_color_mode == aw.saving.ImageColorMode.NONE:
                    self.assertLess(120000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.color_mode.png"))

                elif image_color_mode == aw.saving.ImageColorMode.GRAYSCALE:
                    self.assertLess(80000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.color_mode.png"))

                elif image_color_mode == aw.saving.ImageColorMode.BLACK_AND_WHITE:
                    self.assertGreater(20000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.color_mode.png"))

                #ExEnd

    def test_paper_color(self):

        #ExStart
        #ExFor:ImageSaveOptions
        #ExFor:ImageSaveOptions.paper_color
        #ExSummary:Renders a page of a Word document into an image with transparent or colored background.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.font.name = "Times New Roman"
        builder.font.size = 24
        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")

        builder.insert_image(IMAGE_DIR + "Logo.jpg")

        # Create an "ImageSaveOptions" object which we can pass to the document's "save" method
        # to modify the way in which that method renders the document into an image.
        img_options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)

        # Set the "paper_color" property to a transparent color to apply a transparent
        # background to the document while rendering it to an image.
        img_options.paper_color = drawing.Color.transparent

        doc.save(ARTIFACTS_DIR + "ImageSaveOptions.paper_color.transparent.png", img_options)

        # Set the "paper_color" property to an opaque color to apply that color
        # as the background of the document as we render it to an image.
        img_options.paper_color = drawing.Color.light_coral

        doc.save(ARTIFACTS_DIR + "ImageSaveOptions.paper_color.light_coral.png", img_options)
        #ExEnd

        self.verify_image_contains_transparency(ARTIFACTS_DIR + "ImageSaveOptions.paper_color.transparent.png")
        self.verify_image_contains_transparency(ARTIFACTS_DIR + "ImageSaveOptions.paper_color.light_coral.png")

    def test_pixel_format(self):

        for image_pixel_format in (aw.saving.ImagePixelFormat.FORMAT_1BPP_INDEXED,
                                   aw.saving.ImagePixelFormat.FORMAT_16BPP_RGB_555,
                                   aw.saving.ImagePixelFormat.FORMAT_24BPP_RGB,
                                   aw.saving.ImagePixelFormat.FORMAT_32BPP_RGB,
                                   aw.saving.ImagePixelFormat.FORMAT_48BPP_RGB):
            with self.subTest(image_pixel_format=image_pixel_format):
                #ExStart
                #ExFor:ImagePixelFormat
                #ExFor:ImageSaveOptions.clone
                #ExFor:ImageSaveOptions.pixel_format
                #ExSummary:Shows how to select a bit-per-pixel rate with which to render a document to an image.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.paragraph_format.style = doc.styles.get_by_name("Heading 1")
                builder.writeln("Hello world!")
                builder.insert_image(IMAGE_DIR + "Logo.jpg")

                self.assertLess(20000, os.path.getsize(IMAGE_DIR + "Logo.jpg"))

                # When we save the document as an image, we can pass a SaveOptions object to
                # select a pixel format for the image that the saving operation will generate.
                # Various bit per pixel rates will affect the quality and file size of the generated image.
                image_save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
                image_save_options.pixel_format = image_pixel_format

                # We can clone ImageSaveOptions instances.
                self.assertNotEqual(image_save_options, image_save_options.clone())

                doc.save(ARTIFACTS_DIR + "ImageSaveOptions.pixel_format.png", image_save_options)

                if image_pixel_format == aw.saving.ImagePixelFormat.FORMAT_1BPP_INDEXED:
                    self.assertGreater(10000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.pixel_format.png"))

                elif image_pixel_format == aw.saving.ImagePixelFormat.FORMAT_16BPP_RGB_555:
                    self.assertLess(125000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.pixel_format.png"))

                elif image_pixel_format == aw.saving.ImagePixelFormat.FORMAT_24BPP_RGB:
                    self.assertLess(70000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.pixel_format.png"))

                elif image_pixel_format == aw.saving.ImagePixelFormat.FORMAT_32BPP_RGB:
                    self.assertLess(125000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.pixel_format.png"))

                elif image_pixel_format == aw.saving.ImagePixelFormat.FORMAT_48BPP_RGB:
                    self.assertLess(125000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.pixel_format.png"))

                #ExEnd

    def test_floyd_steinberg_dithering(self):

        #ExStart
        #ExFor:ImageBinarizationMethod
        #ExFor:ImageSaveOptions.threshold_for_floyd_steinberg_dithering
        #ExFor:ImageSaveOptions.tiff_binarization_method
        #ExSummary:Shows how to set the TIFF binarization error threshold when using the Floyd-Steinberg method to render a TIFF image.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.paragraph_format.style = doc.styles.get_by_name("Heading 1")
        builder.writeln("Hello world!")
        builder.insert_image(IMAGE_DIR + "Logo.jpg")

        # When we save the document as a TIFF, we can pass a SaveOptions object to
        # adjust the dithering that Aspose.Words will apply when rendering this image.
        # The default value of the "threshold_for_floyd_steinberg_dithering" property is 128.
        # Higher values tend to produce darker images.
        options = aw.saving.ImageSaveOptions(aw.SaveFormat.TIFF)
        options.tiff_compression = aw.saving.TiffCompression.CCITT3
        options.tiff_binarization_method = aw.saving.ImageBinarizationMethod.FLOYD_STEINBERG_DITHERING
        options.threshold_for_floyd_steinberg_dithering = 240

        doc.save(ARTIFACTS_DIR + "ImageSaveOptions.floyd_steinberg_dithering.tiff", options)
        #ExEnd

        self.verify_image(816, 1056, filename=ARTIFACTS_DIR + "ImageSaveOptions.floyd_steinberg_dithering.tiff")

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
        builder = aw.DocumentBuilder(doc)

        builder.paragraph_format.style = doc.styles.get_by_name("Heading 1")
        builder.writeln("Hello world!")
        builder.insert_image(IMAGE_DIR + "Logo.jpg")

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

        doc.save(ARTIFACTS_DIR + "ImageSaveOptions.edit_image.png", options)
        #ExEnd

        self.verify_image(817, 1057, filename=ARTIFACTS_DIR + "ImageSaveOptions.edit_image.png")

    def test_jpeg_quality(self):

        #ExStart
        #ExFor:Document.save(str,SaveOptions)
        #ExFor:FixedPageSaveOptions.jpeg_quality
        #ExFor:ImageSaveOptions
        #ExFor:ImageSaveOptions.__init__
        #ExFor:ImageSaveOptions.jpeg_quality
        #ExSummary:Shows how to configure compression while saving a document as a JPEG.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.insert_image(IMAGE_DIR + "Logo.jpg")

        # Create an "ImageSaveOptions" object which we can pass to the document's "save" method
        # to modify the way in which that method renders the document into an image.
        image_options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)

        # Set the "jpeg_quality" property to "10" to use stronger compression when rendering the document.
        # This will reduce the file size of the document, but the image will display more prominent compression artifacts.
        image_options.jpeg_quality = 10

        doc.save(ARTIFACTS_DIR + "ImageSaveOptions.jpeg_quality.high_compression.jpg", image_options)

        self.assertGreater(20000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.jpeg_quality.high_compression.jpg"))

        # Set the "jpeg_quality" property to "100" to use weaker compression when rending the document.
        # This will improve the quality of the image at the cost of an increased file size.
        image_options.jpeg_quality = 100

        doc.save(ARTIFACTS_DIR + "ImageSaveOptions.jpeg_quality.high_quality.jpg", image_options)

        self.assertLess(60000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.jpeg_quality.high_quality.jpg"))
        #ExEnd

    def test_save_to_tiff_default(self):

        doc = aw.Document(MY_DIR + "Rendering.docx")
        doc.save(ARTIFACTS_DIR + "ImageSaveOptions.save_to_tiff_default.tiff")

    def test_tiff_image_compression(self):

        for tiff_compression in (aw.saving.TiffCompression.NONE,
                                 aw.saving.TiffCompression.RLE,
                                 aw.saving.TiffCompression.LZW,
                                 aw.saving.TiffCompression.CCITT3,
                                 aw.saving.TiffCompression.CCITT4):
            with self.subTest(tiff_compression=tiff_compression):
                #ExStart
                #ExFor:TiffCompression
                #ExFor:ImageSaveOptions.tiff_compression
                #ExSummary:Shows how to select the compression scheme to apply to a document that we convert into a TIFF image.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.insert_image(IMAGE_DIR + "Logo.jpg")

                # Create an "ImageSaveOptions" object which we can pass to the document's "save" method
                # to modify the way in which that method renders the document into an image.
                options = aw.saving.ImageSaveOptions(aw.SaveFormat.TIFF)

                # Set the "tiff_compression" property to "TiffCompression.NONE" to apply no compression while saving,
                # which may result in a very large output file.
                # Set the "tiff_compression" property to "TiffCompression.RLE" to apply RLE compression
                # Set the "tiff_compression" property to "TiffCompression.LZW" to apply LZW compression.
                # Set the "tiff_compression" property to "TiffCompression.CCITT3" to apply CCITT3 compression.
                # Set the "tiff_compression" property to "TiffCompression.CCITT4" to apply CCITT4 compression.
                options.tiff_compression = tiff_compression

                doc.save(ARTIFACTS_DIR + "ImageSaveOptions.tiff_image_compression.tiff", options)

                if tiff_compression == aw.saving.TiffCompression.NONE:
                    self.assertLess(3000000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.tiff_image_compression.tiff"))

                elif tiff_compression == aw.saving.TiffCompression.RLE:
                    self.assertLess(6000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.tiff_image_compression.tiff"))

                elif tiff_compression == aw.saving.TiffCompression.LZW:
                    self.assertLess(200000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.tiff_image_compression.tiff"))

                elif tiff_compression == aw.saving.TiffCompression.CCITT3:
                    self.assertGreater(90000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.tiff_image_compression.tiff"))

                elif tiff_compression == aw.saving.TiffCompression.CCITT4:
                    self.assertGreater(20000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.tiff_image_compression.tiff"))

                #ExEnd

    def test_resolution(self):

        #ExStart
        #ExFor:ImageSaveOptions
        #ExFor:ImageSaveOptions.resolution
        #ExSummary:Shows how to specify a resolution while rendering a document to PNG.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.font.name = "Times New Roman"
        builder.font.size = 24
        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")

        builder.insert_image(IMAGE_DIR + "Logo.jpg")

        # Create an "ImageSaveOptions" object which we can pass to the document's "save" method
        # to modify the way in which that method renders the document into an image.
        options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)

        # Set the "resolution" property to "72" to render the document in 72dpi.
        options.vertical_resolution = 72
        options.horizontal_resolution = 72

        doc.save(ARTIFACTS_DIR + "ImageSaveOptions.resolution.72dpi.png", options)

        self.assertGreater(120000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.resolution.72dpi.png"))

        image = drawing.Image.from_file(ARTIFACTS_DIR + "ImageSaveOptions.resolution.72dpi.png")

        self.assertEqual(612, image.width)
        self.assertEqual(792, image.height)

        # Set the "resolution" property to "300" to render the document in 300dpi.
        options.vertical_resolution = 300
        options.horizontal_resolution = 300

        doc.save(ARTIFACTS_DIR + "ImageSaveOptions.resolution.300dpi.png", options)

        self.assertLess(700000, os.path.getsize(ARTIFACTS_DIR + "ImageSaveOptions.resolution.300dpi.png"))

        image = drawing.Image.from_file(ARTIFACTS_DIR + "ImageSaveOptions.resolution.300dpi.png")

        self.assertEqual(2550, image.width)
        self.assertEqual(3300, image.height)

        #ExEnd

    def test_export_various_page_ranges(self):

        #ExStart
        #ExFor:PageSet.__init__(List[PageRange])
        #ExFor:PageRange.__init__(int,int)
        #ExFor:ImageSaveOptions.page_set
        #ExSummary:Shows how to extract pages based on exact page ranges.
        doc = aw.Document(MY_DIR + "Images.docx")

        image_options = aw.saving.ImageSaveOptions(aw.SaveFormat.TIFF)
        page_set = aw.saving.PageSet([
            aw.saving.PageRange(1, 1),
            aw.saving.PageRange(2, 3),
            aw.saving.PageRange(1, 3),
            aw.saving.PageRange(2, 4),
            aw.saving.PageRange(1, 1)])

        image_options.page_set = page_set
        doc.save(ARTIFACTS_DIR + "ImageSaveOptions.export_various_page_ranges.tiff", image_options)
        #ExEnd

    def test_render_ink_object(self):

        #ExStart
        #ExFor:SaveOptions.iml_rendering_mode
        #ExFor:ImlRenderingMode
        #ExSummary:Shows how to render Ink object.
        doc = aw.Document(MY_DIR + "Ink object.docx")

        # Set 'ImlRenderingMode.INK_ML' ignores fall-back shape of ink (InkML) object and renders InkML itself.
        # If the rendering result is unsatisfactory,
        # please use 'ImlRenderingMode.FALLBACK' to get a result similar to previous versions.
        save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.JPEG)
        save_options.iml_rendering_mode = aw.saving.ImlRenderingMode.INK_ML

        doc.save(ARTIFACTS_DIR + "ImageSaveOptions.render_ink_object.jpeg", save_options)
        #ExEnd
