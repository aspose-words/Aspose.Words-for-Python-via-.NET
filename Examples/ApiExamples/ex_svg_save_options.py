import unittest
import io
import os

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

class ExSvgSaveOptions(ApiExampleBase):

    def test_save_like_image(self):

        #ExStart
        #ExFor:SvgSaveOptions.FitToViewPort
        #ExFor:SvgSaveOptions.ShowPageBorder
        #ExFor:SvgSaveOptions.TextOutputMode
        #ExFor:SvgTextOutputMode
        #ExSummary:Shows how to mimic the properties of images when converting a .docx document to .svg.
        doc = aw.Document(MY_DIR + "Document.docx")

        # Configure the SvgSaveOptions object to save with no page borders or selectable text.
        options = aw.saving.SvgSaveOptions()
        options.fit_to_view_port = True
        options.show_page_border = False
        options.text_output_mode = aw.saving.SvgTextOutputMode.USE_PLACED_GLYPHS

        doc.save(ARTIFACTS_DIR + "SvgSaveOptions.SaveLikeImage.svg", options)
        #ExEnd

    #ExStart
    #ExFor:SvgSaveOptions
    #ExFor:SvgSaveOptions.ExportEmbeddedImages
    #ExFor:SvgSaveOptions.ResourceSavingCallback
    #ExFor:SvgSaveOptions.ResourcesFolder
    #ExFor:SvgSaveOptions.ResourcesFolderAlias
    #ExFor:SvgSaveOptions.SaveFormat
    #ExSummary:Shows how to manipulate and print the URIs of linked resources created while converting a document to .svg.
    #def test_svg_resource_folder(self):

    #    doc = aw.Document(MY_DIR + "Rendering.docx")

    #    options = aw.saving.SvgSaveOptions()
    #    options.save_format = aw.saving.SaveFormat.SVG
    #    options.export_embedded_images = False
    #    options.resources_folder = ARTIFACTS_DIR + "SvgResourceFolder"
    #    options.resources_folder_alias = ARTIFACTS_DIR + "SvgResourceFolderAlias",
    #    options.show_page_border = False
    #    options.resource_saving_callback = ExSvgSaveOptions.ResourceUriPrinter()

    #    os.mkdir(options.resources_folder_alias)

    #    doc.save(ARTIFACTS_DIR + "SvgSaveOptions.SvgResourceFolder.svg", options)

    #class ResourceUriPrinter(aw.saving.IResourceSavingCallback):
    #    """Counts and prints URIs of resources contained by as they are converted to .svg."""

    #    def __init__(self):
    #        self.saved_resource_count = 0

    #    def resource_saving(self, args: aw.saving.ResourceSavingArgs):
    #        self.saved_resource_count += 1
    #        print(f"Resource #{self.saved_resource_count} \"{args.resource_file_name}\"")
    #        print("\t" + args.resource_file_uri)

    #ExEnd
