# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
import aspose.words.drawing
import aspose.words.saving
import test_util
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, IMAGE_DIR, MY_DIR

class ExRtfSaveOptions(ApiExampleBase):

    def test_save_images_as_wmf(self):
        for save_images_as_wmf in [False, True]:
            #ExStart
            #ExFor:RtfSaveOptions.save_images_as_wmf
            #ExSummary:Shows how to convert all images in a document to the Windows Metafile format as we save the document as an RTF.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.writeln('Jpeg image:')
            image_shape = builder.insert_image(file_name=IMAGE_DIR + 'Logo.jpg')
            self.assertEqual(aw.drawing.ImageType.JPEG, image_shape.image_data.image_type)
            builder.insert_paragraph()
            builder.writeln('Png image:')
            image_shape = builder.insert_image(file_name=IMAGE_DIR + 'Transparent background logo.png')
            self.assertEqual(aw.drawing.ImageType.PNG, image_shape.image_data.image_type)
            # Create an "RtfSaveOptions" object to pass to the document's "Save" method to modify how we save it to an RTF.
            rtf_save_options = aw.saving.RtfSaveOptions()
            # Set the "SaveImagesAsWmf" property to "true" to convert all images in the document to WMF as we save it to RTF.
            # Doing so will help readers such as WordPad to read our document.
            # Set the "SaveImagesAsWmf" property to "false" to preserve the original format of all images in the document
            # as we save it to RTF. This will preserve the quality of the images at the cost of compatibility with older RTF readers.
            rtf_save_options.save_images_as_wmf = save_images_as_wmf
            doc.save(file_name=ARTIFACTS_DIR + 'RtfSaveOptions.SaveImagesAsWmf.rtf', save_options=rtf_save_options)
            doc = aw.Document(file_name=ARTIFACTS_DIR + 'RtfSaveOptions.SaveImagesAsWmf.rtf')
            shapes = doc.get_child_nodes(aw.NodeType.SHAPE, True)
            if save_images_as_wmf:
                self.assertEqual(aw.drawing.ImageType.WMF, shapes[0].as_shape().image_data.image_type)
                self.assertEqual(aw.drawing.ImageType.WMF, shapes[1].as_shape().image_data.image_type)
            else:
                self.assertEqual(aw.drawing.ImageType.JPEG, shapes[0].as_shape().image_data.image_type)
                self.assertEqual(aw.drawing.ImageType.PNG, shapes[1].as_shape().image_data.image_type)
            #ExEnd

    def test_export_images(self):
        for export_images_for_old_readers in (False, True):
            with self.subTest(export_images_for_old_readers=export_images_for_old_readers):
                #ExStart
                #ExFor:RtfSaveOptions
                #ExFor:RtfSaveOptions.export_compact_size
                #ExFor:RtfSaveOptions.export_images_for_old_readers
                #ExFor:RtfSaveOptions.save_format
                #ExSummary:Shows how to save a document to .rtf with custom options.
                doc = aw.Document(MY_DIR + 'Rendering.docx')
                # Create an "RtfSaveOptions" object to pass to the document's "save" method to modify how we save it to an RTF.
                options = aw.saving.RtfSaveOptions()
                self.assertEqual(aw.SaveFormat.RTF, options.save_format)
                # Set the "export_compact_size" property to "True" to
                # reduce the saved document's size at the cost of right-to-left text compatibility.
                options.export_compact_size = True
                # Set the "export_images_for_old_readers" property to "True" to use extra keywords to ensure that our document is
                # compatible with pre-Microsoft Word 97 readers and WordPad.
                # Set the "export_images_for_old_readers" property to "False" to reduce the size of the document,
                # but prevent old readers from being able to read any non-metafile or BMP images that the document may contain.
                options.export_images_for_old_readers = export_images_for_old_readers
                doc.save(ARTIFACTS_DIR + 'RtfSaveOptions.export_images.rtf', options)
                #ExEnd
                with open(ARTIFACTS_DIR + 'RtfSaveOptions.export_images.rtf', 'rb') as file:
                    data = file.read().decode('utf-8')
                    if export_images_for_old_readers:
                        self.assertIn('nonshppict', data)
                        self.assertIn('shprslt', data)
                    else:
                        self.assertNotIn('nonshppict', data)
                        self.assertNotIn('shprslt', data)