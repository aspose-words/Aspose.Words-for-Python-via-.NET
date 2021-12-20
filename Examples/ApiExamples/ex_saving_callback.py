# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import os

import aspose.words as aw

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, IMAGE_DIR

#class ExSavingCallback(ApiExampleBase):

#    def test_check_that_all_methods_are_present(self):

#        html_fixed_save_options = aw.saving.HtmlFixedSaveOptions()
#        html_fixed_save_options.page_saving_callback = ExSavingCallback.CustomFileNamePageSavingCallback()

#        image_save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
#        image_save_options.page_saving_callback = ExSavingCallback.CustomFileNamePageSavingCallback()

#        pdf_save_options = aw.saving.PdfSaveOptions()
#        pdf_save_options.page_saving_callback = ExSavingCallback.CustomFileNamePageSavingCallback()

#        ps_save_options = aw.saving.PsSaveOptions()
#        ps_save_options.page_saving_callback = ExSavingCallback.CustomFileNamePageSavingCallback()

#        svg_save_options = aw.saving.SvgSaveOptions()
#        svg_save_options.page_saving_callback = ExSavingCallback.CustomFileNamePageSavingCallback()

#        xaml_fixed_save_options = aw.saving.XamlFixedSaveOptions()
#        xaml_fixed_save_options.page_saving_callback = ExSavingCallback.CustomFileNamePageSavingCallback()

#        xps_save_options = aw.saving.XpsSaveOptions()
#        xps_save_options.page_saving_callback = ExSavingCallback.CustomFileNamePageSavingCallback()

#    #ExStart
#    #ExFor:IPageSavingCallback
#    #ExFor:IPageSavingCallback.page_saving(PageSavingArgs)
#    #ExFor:PageSavingArgs
#    #ExFor:PageSavingArgs.page_file_name
#    #ExFor:PageSavingArgs.keep_page_stream_open
#    #ExFor:PageSavingArgs.page_index
#    #ExFor:PageSavingArgs.page_stream
#    #ExFor:FixedPageSaveOptions.page_saving_callback
#    #ExSummary:Shows how to use a callback to save a document to HTML page by page.
#    def test_page_file_names(self):

#        doc = aw.Document()
#        builder = aw.DocumentBuilder(doc)

#        builder.writeln("Page 1.")
#        builder.insert_break(aw.break_type.PAGE_BREAK)
#        builder.writeln("Page 2.")
#        builder.insert_image(IMAGE_DIR + "Logo.jpg")
#        builder.insert_break(aw.break_type.PAGE_BREAK)
#        builder.writeln("Page 3.")

#        # Create an "HtmlFixedSaveOptions" object, which we can pass to the document's "save" method
#        # to modify how we convert the document to HTML.
#        html_fixed_save_options = aw.saving.HtmlFixedSaveOptions()

#        # We will save each page in this document to a separate HTML file in the local file system.
#        # Set a callback that allows us to name each output HTML document.
#        html_fixed_save_options.page_saving_callback = ExSavingCallback.CustomFileNamePageSavingCallback()

#        doc.save(ARTIFACTS_DIR + "SavingCallback.page_file_names.html", html_fixed_save_options)

#        file_paths = [path for path in os.listdir(ARTIFACTS_DIR) if path.startswith(ARTIFACTS_DIR + "SavingCallback.page_file_names.Page_")]

#        self.assertEqual(3, len(file_paths))

#    class CustomFileNamePageSavingCallback(aw.saving.IPageSavingCallback):
#        """Saves all pages to a file and directory specified within."""

#        def page_saving(args: aw.saving.PageSavingArgs):

#            out_file_name = f"{ARTIFACTS_DIR}SavingCallback.page_file_names.Page_{args.page_index}.html"

#            # Below are two ways of specifying where Aspose.Words will save each page of the document.
#            # 1 -  Set a filename for the output page file:
#            args.page_file_name = out_file_name

#            # 2 -  Create a custom stream for the output page file:
#            args.page_stream = open(out_file_name, 'wb')

#            self.assertFalse(args.keep_page_stream_open)

#    #ExEnd

#    #ExStart
#    #ExFor:DocumentPartSavingArgs
#    #ExFor:DocumentPartSavingArgs.document
#    #ExFor:DocumentPartSavingArgs.document_part_file_name
#    #ExFor:DocumentPartSavingArgs.document_part_stream
#    #ExFor:DocumentPartSavingArgs.keep_document_part_stream_open
#    #ExFor:IDocumentPartSavingCallback
#    #ExFor:IDocumentPartSavingCallback.document_part_saving(DocumentPartSavingArgs)
#    #ExFor:IImageSavingCallback
#    #ExFor:IImageSavingCallback.image_saving
#    #ExFor:ImageSavingArgs
#    #ExFor:ImageSavingArgs.image_file_name
#    #ExFor:HtmlSaveOptions
#    #ExFor:HtmlSaveOptions.document_part_saving_callback
#    #ExFor:HtmlSaveOptions.image_saving_callback
#    #ExSummary:Shows how to split a document into parts and save them.
#    def test_document_parts_file_names(self):

#        doc = aw.Document(MY_DIR + "Rendering.docx")
#        out_file_name = "SavingCallback.document_parts_file_names.html"

#        # Create an "HtmlFixedSaveOptions" object, which we can pass to the document's "save" method
#        # to modify how we convert the document to HTML.
#        options = aw.saving.HtmlSaveOptions()

#        # If we save the document normally, there will be one output HTML
#        # document with all the source document's contents.
#        # Set the "document_split_criteria" property to "DocumentSplitCriteria.SECTION_BREAK" to
#        # save our document to multiple HTML files: one for each section.
#        options.document_split_criteria = aw.saving.DocumentSplitCriteria.SECTION_BREAK

#        # Assign a custom callback to the "document_part_saving_callback" property to alter the document part saving logic.
#        options.document_part_saving_callback = ExSavingCallback.SavedDocumentPartRename(out_file_name, options.document_split_criteria)

#        # If we convert a document that contains images into html, we will end up with one html file which links to several images.
#        # Each image will be in the form of a file in the local file system.
#        # There is also a callback that can customize the name and file system location of each image.
#        options.image_saving_callback = aw.saving.SavedImageRename(out_file_name)

#        doc.save(ARTIFACTS_DIR + out_file_name, options)

#    class SavedDocumentPartRename(aw.saving.IDocumentPartSavingCallback):
#        """Sets custom filenames for output documents that the saving operation splits a document into."""

#        def __init__(self, out_file_name: str, document_split_criteria: aw.saving.DocumentSplitCriteria):

#            self.out_file_name = out_file_name
#            self.document_split_criteria = document_split_criteria

#        def document_part_saving(self, args: aw.saving.DocumentPartSavingArgs):

#            # We can access the entire source document via the "Document" property.
#            self.assertTrue(args.document.original_file_name.endswith("Rendering.docx"))

#            if self.document_split_criteria == aw.saving.DocumentSplitCriteria.PAGE_BREAK:
#                part_type = "Page"
#            elif self.document_split_criteria == aw.saving.DocumentSplitCriteria.COLUMN_BREAK:
#                part_type = "Column"
#            elif self.document_split_criteria == aw.saving.DocumentSplitCriteria.SECTION_BREAK:
#                part_type = "Section"
#            elif self.document_split_criteria == aw.saving.DocumentSplitCriteria.HEADING_PARAGRAPH:
#                part_type = "Paragraph from heading"

#            part_file_name = f"{self.out_file_name} part {++self.count}, of type {part_type}{os.path.splitext(args.document_part_file_name)[1]}"

#            # Below are two ways of specifying where Aspose.Words will save each part of the document.
#            # 1 -  Set a filename for the output part file:
#            args.document_part_file_name = part_file_name

#            # 2 -  Create a custom stream for the output part file:
#            args.document_part_stream = open(ARTIFACTS_DIR + part_file_name, 'wb')

#            self.assertTrue(args.document_part_stream.can_write)
#            self.assertFalse(args.keep_document_part_stream_open)

#    class SavedImageRename(aw.saving.IImageSavingCallback):
#        """Sets custom filenames for image files that an HTML conversion creates."""

#        def __init__(self, out_file_name: str):

#            self.out_file_name = out_file_name
#            self.count = 0

#        def image_saving(args: aw.saving.ImageSavingArgs):

#            self.count += 1
#            image_file_name = f"{self.out_file_name} shape {self.count}, of type {args.current_shape.shape_type}{os.path.splitext(args.image_file_name)[1]}"

#            # Below are two ways of specifying where Aspose.Words will save each part of the document.
#            # 1 -  Set a filename for the output image file:
#            args.image_file_name = image_file_name

#            # 2 -  Create a custom stream for the output image file:
#            args.image_stream = open(ARTIFACTS_DIR + image_file_name, 'wb')

#            self.assertTrue(args.image_stream.can_write)
#            self.assertTrue(args.is_image_available)
#            self.assertFalse(args.keep_image_stream_open)

#    #ExEnd

#    #ExStart
#    #ExFor:CssSavingArgs
#    #ExFor:CssSavingArgs.css_stream
#    #ExFor:CssSavingArgs.document
#    #ExFor:CssSavingArgs.is_export_needed
#    #ExFor:CssSavingArgs.keep_css_stream_open
#    #ExFor:CssStyleSheetType
#    #ExFor:HtmlSaveOptions.css_saving_callback
#    #ExFor:HtmlSaveOptions.css_style_sheet_file_name
#    #ExFor:HtmlSaveOptions.css_style_sheet_type
#    #ExFor:ICssSavingCallback
#    #ExFor:ICssSavingCallback.css_saving(CssSavingArgs)
#    #ExSummary:Shows how to work with CSS stylesheets that an HTML conversion creates.
#    def test_external_css_filenames(self):

#        doc = aw.Document(MY_DIR + "Rendering.docx")

#        # Create an "HtmlFixedSaveOptions" object, which we can pass to the document's "save" method
#        # to modify how we convert the document to HTML.
#        options = aw.saving.HtmlSaveOptions()

#        # Set the "css_style_sheet_type" property to "CssStyleSheetType.EXTERNAL" to
#        # accompany a saved HTML document with an external CSS stylesheet file.
#        options.css_style_sheet_type = aw.saving.CssStyleSheetType.EXTERNAL

#        # Below are two ways of specifying directories and filenames for output CSS stylesheets.
#        # 1 -  Use the "CssStyleSheetFileName" property to assign a filename to our stylesheet:
#        options.css_style_sheet_file_name = ARTIFACTS_DIR + "SavingCallback.external_css_filenames.css"

#        # 2 -  Use a custom callback to name our stylesheet:
#        options.css_saving_callback = ExSavingCallback.CustomCssSavingCallback(ARTIFACTS_DIR + "SavingCallback.external_css_filenames.css", True, False)

#        doc.save(ARTIFACTS_DIR + "SavingCallback.external_css_filenames.html", options)

#    class CustomCssSavingCallback(aw.saving.ICssSavingCallback):
#        """Sets a custom filename, along with other parameters for an external CSS stylesheet."""

#        def __init__(self, css_doc_filename: str, is_export_needed: bool, keep_css_stream_open: bool):

#            self.css_text_file_name = css_doc_filename
#            self.is_export_needed = is_export_needed
#            self.keep_css_stream_open = keep_css_stream_open

#        def css_saving(self, args: aw.saving.CssSavingArgs):

#            # We can access the entire source document via the "Document" property.
#            self.assertTrue(args.document.original_file_name.ends_with("Rendering.docx"))

#            args.css_stream = open(self.css_text_file_name, 'wb')
#            args.is_export_needed = self.is_export_needed
#            args.keep_css_stream_open = self.keep_css_stream_open

#            self.assertTrue(args.css_stream.can_write)

#    #ExEnd
