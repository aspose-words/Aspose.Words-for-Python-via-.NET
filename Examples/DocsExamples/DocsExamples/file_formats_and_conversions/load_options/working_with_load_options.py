import aspose.words as aw
import unittest
import sys
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR, IMAGES_DIR

class WorkingWithLoadOptions(DocsExamplesBase):

    def test_update_dirty_fields(self):

        #ExStart:UpdateDirtyFields
        #GistId:365214c8b2e8c065166447871a1499aa
        load_options = aw.loading.LoadOptions()
        load_options.update_dirty_fields = True

        doc = aw.Document(MY_DIR + "Dirty field.docx", load_options)

        doc.save(ARTIFACTS_DIR + "WorkingWithLoadOptions.update_dirty_fields.docx")
        #ExEnd:UpdateDirtyFields

    def test_load_encrypted_document(self):

        #ExStart:LoadSaveEncryptedDocument
        #GistId:6548546f98bd830e363bbb567b114850
        #ExStart:OpenEncryptedDocument
        #GistId:41c71acaf4924abe47f4bc2ff2c87d6a
        doc = aw.Document(MY_DIR + "Encrypted.docx", aw.loading.LoadOptions("docPassword"))
        #ExEnd:OpenEncryptedDocument

        doc.save(ARTIFACTS_DIR + "WorkingWithLoadOptions.load_and_save_encrypted_odt.odt", aw.saving.OdtSaveOptions("newPassword"))
        #ExEnd:LoadSaveEncryptedDocument

    def test_load_encrypted_document_without_password(self):
        #ExStart:LoadEncryptedDocumentWithoutPassword
        #GistId:6548546f98bd830e363bbb567b114850
        # We will not be able to open this document with Microsoft Word or
        # Aspose.Words without providing the correct password.
        with self.assertRaises(Exception):
            aw.Document(MY_DIR + "Encrypted.docx")
        #ExEnd:LoadEncryptedDocumentWithoutPassword

    def test_convert_shape_to_office_math(self):

        #ExStart:ConvertShapeToOfficeMath
        #GistId:664d989412b46c6a03619182c5a4b9cd
        load_options = aw.loading.LoadOptions()
        load_options.convert_shape_to_office_math = True

        doc = aw.Document(MY_DIR + "Office math.docx", load_options)

        doc.save(ARTIFACTS_DIR + "WorkingWithLoadOptions.convert_shape_to_office_math.docx", aw.SaveFormat.DOCX)
        #ExEnd:ConvertShapeToOfficeMath

    def test_set_ms_word_version(self):

        #ExStart:SetMsWordVersion
        #GistId:41c71acaf4924abe47f4bc2ff2c87d6a
        # Create a new LoadOptions object, which will load documents according to MS Word 2019 specification by default
        # and change the loading version to Microsoft Word 2010.
        load_options = aw.loading.LoadOptions()
        load_options.msw_version = aw.settings.MsWordVersion.WORD2010

        doc = aw.Document(MY_DIR + "Document.docx", load_options)

        doc.save(ARTIFACTS_DIR + "WorkingWithLoadOptions.set_ms_word_version.docx")
        #ExEnd:SetMsWordVersion

    def test_temp_folder(self):

        #ExStart:TempFolder
        #GistId:41c71acaf4924abe47f4bc2ff2c87d6a
        load_options = aw.loading.LoadOptions()
        load_options.temp_folder = ARTIFACTS_DIR

        doc = aw.Document(MY_DIR + "Document.docx", load_options)
        #ExEnd:TempFolder

    @unittest.skipUnless(sys.platform.startswith('win'), 'requires windows')
    def test_load_with_encoding(self):

        #ExStart:LoadWithEncoding
        #GistId:41c71acaf4924abe47f4bc2ff2c87d6a
        load_options = aw.loading.LoadOptions()
        load_options.encoding = "utf-7"

        doc = aw.Document(MY_DIR + "Encoded in UTF-7.txt", load_options)
        #ExEnd:LoadWithEncoding

    def test_skip_pdf_images(self):

        #ExStart:SkipPdfImages
        load_options = aw.loading.PdfLoadOptions()
        load_options.skip_pdf_images = True

        doc = aw.Document(MY_DIR + "Pdf Document.pdf", load_options)
        #ExEnd:SkipPdfImages

    def test_convert_metafiles_to_png(self):

        #ExStart:ConvertMetafilesToPng
        load_options = aw.loading.LoadOptions()
        load_options.convert_metafiles_to_png = True

        doc = aw.Document(MY_DIR + "WMF with image.docx", load_options)
        #ExEnd:ConvertMetafilesToPng

    def test_load_chm(self):

        #ExStart:LoadChm
        load_options = aw.loading.LoadOptions()
        load_options.encoding = "windows-1251"

        doc = aw.Document(MY_DIR + "HTML help.chm", load_options)
        #ExEnd:LoadChm

    def test_warning_callback(self):

        #ExStart:WarningCallback
        #GistId:41c71acaf4924abe47f4bc2ff2c87d6a
        load_options = aw.loading.LoadOptions()
        load_options.warning_callback = DocumentLoadingWarningCallback()

        doc = aw.Document(MY_DIR + "Document.docx", load_options)
        #ExEnd:WarningCallback

    def test_resource_loading_callback(self):

        #ExStart:ResourceLoadingCallback
        #GistId:41c71acaf4924abe47f4bc2ff2c87d6a
        load_options = aw.loading.LoadOptions()
        load_options.resource_loading_callback = HtmlLinkedResourceLoadingCallback()

        # When we open an Html document, external resources such as references to CSS stylesheet files
        # and external images will be handled customarily by the loading callback as the document is loaded.
        doc = aw.Document(MY_DIR + "Images.html", load_options)

        doc.save(ARTIFACTS_DIR + "WorkingWithLoadOptions.resource_loading_callback.pdf")
        #ExEnd:ResourceLoadingCallback


#ExStart:IWarningCallback
#GistId:41c71acaf4924abe47f4bc2ff2c87d6a
class DocumentLoadingWarningCallback(aw.IWarningCallback):

    def warning(self, info: aw.WarningInfo):

        # Prints warnings and their details as they arise during document loading.
        print(f"WARNING: {info.warning_type}, source: {info.source}")
        print(f"	Description: {info.description}")
#ExEnd:IWarningCallback


#ExStart:IResourceLoadingCallback
#GistId:41c71acaf4924abe47f4bc2ff2c87d6a
class HtmlLinkedResourceLoadingCallback(aw.loading.IResourceLoadingCallback):

    def resource_loading(self, args: aw.loading.ResourceLoadingArgs):

        if args.resource_type == aw.loading.ResourceType.CSS_STYLE_SHEET:
            print(f"External CSS Stylesheet found upon loading: {args.original_uri}")

            # CSS file will don't used in the document.
            return aw.loading.ResourceLoadingAction.SKIP

        if args.resource_type == aw.loading.ResourceType.IMAGE:
            # Replaces all images with a substitute.
            with open(IMAGES_DIR + "Logo.jpg", "rb") as file:
                args.set_data(file.read())

            # New images will be used instead of presented in the document.
            return aw.loading.ResourceLoadingAction.USER_PROVIDED

        if args.resource_type == aw.loading.ResourceType.DOCUMENT:
            print(f"External document found upon loading: {args.original_uri}")

            # Will be used as usual.
            return aw.loading.ResourceLoadingAction.DEFAULT

        raise RuntimeError("Unexpected ResourceType value.")
#ExEnd:IResourceLoadingCallback
