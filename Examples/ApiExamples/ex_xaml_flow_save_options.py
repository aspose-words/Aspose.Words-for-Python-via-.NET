# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import os

import aspose.words as aw

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExXamlFlowSaveOptions(ApiExampleBase):

    #ExStart
    #ExFor:XamlFlowSaveOptions
    #ExFor:XamlFlowSaveOptions.__init__
    #ExFor:XamlFlowSaveOptions.__init__(SaveFormat)
    #ExFor:XamlFlowSaveOptions.image_saving_callback
    #ExFor:XamlFlowSaveOptions.images_folder
    #ExFor:XamlFlowSaveOptions.images_folder_alias
    #ExFor:XamlFlowSaveOptions.save_format
    #ExSummary:Shows how to print the filenames of linked images created while converting a document to flow-form .xaml.
    def test_image_folder(self):

        doc = aw.Document(MY_DIR + "Rendering.docx")

        callback = ExXamlFlowSaveOptions.ImageUriPrinter(ARTIFACTS_DIR + "XamlFlowImageFolderAlias")

        # Create a "XamlFlowSaveOptions" object, which we can pass to the document's "save" method
        # to modify how we save the document to the XAML save format.
        options = aw.saving.XamlFlowSaveOptions()

        self.assertEqual(aw.SaveFormat.XAML_FLOW, options.save_format)

        # Use the "images_folder" property to assign a folder in the local file system into which
        # Aspose.Words will save all the document's linked images.
        options.images_folder = ARTIFACTS_DIR + "XamlFlowImageFolder"

        # Use the "images_folder_alias" property to use this folder
        # when constructing image URIs instead of the images folder's name.
        options.images_folder_alias = ARTIFACTS_DIR + "XamlFlowImageFolderAlias"

        options.image_saving_callback = callback

        # A folder specified by "images_folder_alias" will need to contain the resources instead of "images_folder".
        # We must ensure the folder exists before the callback's streams can put their resources into it.
        os.makedirs(options.images_folder_alias)

        doc.save(ARTIFACTS_DIR + "XamlFlowSaveOptions.image_folder.xaml", options)

        for resource in callback.Resources:
            print(f"{callback.images_folder_alias}/{resource}")

        self._test_image_folder(callback) #ExSkip

    class ImageUriPrinter(aw.saving.IImageSavingCallback):
        """Counts and prints filenames of images while their parent document is converted to flow-form .xaml."""

        def __init__(self, images_folder_alias: str):

            self.images_folder_alias = images_folder_alias
            self.resources = [] # type: List[str]

        def image_saving(self, args: aw.saving.ImageSavingArgs):

            self.resources.add(args.image_file_name)

            # If we specified an image folder alias, we would also need
            # to redirect each stream to put its image in the alias folder.
            args.image_stream = open(f"{images_folder_alias}/{args.image_file_name}", "wb")
            args.keep_image_stream_open = False

    #ExEnd

    def _test_image_folder(self, callback: ExXamlFlowSaveOptions.ImageUriPrinter):

        self.assertEqual(9, len(callback.resources))
        for resource in callback.resources:
            self.assertTrue(os.path.exists(f"{callback.images_folder_alias}/{resource}"))
