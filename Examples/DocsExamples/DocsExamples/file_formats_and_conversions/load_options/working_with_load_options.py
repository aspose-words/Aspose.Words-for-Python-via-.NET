import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class WorkingWithLoadOptions(DocsExamplesBase):

    def test_update_dirty_fields(self):

        #ExStart:UpdateDirtyFields
        load_options = aw.loading.LoadOptions()
        load_options.update_dirty_fields = True

        doc = aw.Document(MY_DIR + "Dirty field.docx", load_options)

        doc.save(ARTIFACTS_DIR + "WorkingWithLoadOptions.update_dirty_fields.docx")
        #ExEnd:UpdateDirtyFields

    def test_load_encrypted_document(self):

        #ExStart:LoadSaveEncryptedDoc
        #ExStart:OpenEncryptedDocument
        doc = aw.Document(MY_DIR + "Encrypted.docx", aw.loading.LoadOptions("docPassword"))
        #ExEnd:OpenEncryptedDocument

        doc.save(ARTIFACTS_DIR + "WorkingWithLoadOptions.load_and_save_encrypted_odt.odt", aw.saving.OdtSaveOptions("newPassword"))
        #ExEnd:LoadSaveEncryptedDoc

    def test_convert_shape_to_office_math(self):

        #ExStart:ConvertShapeToOfficeMath
        load_options = aw.loading.LoadOptions()
        load_options.convert_shape_to_office_math = True

        doc = aw.Document(MY_DIR + "Office math.docx", load_options)

        doc.save(ARTIFACTS_DIR + "WorkingWithLoadOptions.convert_shape_to_office_math.docx", aw.SaveFormat.DOCX)
        #ExEnd:ConvertShapeToOfficeMath

    def test_set_ms_word_version(self):

        #ExStart:SetMSWordVersion
        # Create a new LoadOptions object, which will load documents according to MS Word 2019 specification by default
        # and change the loading version to Microsoft Word 2010.
        load_options = aw.loading.LoadOptions()
        load_options.msw_version = aw.settings.MsWordVersion.WORD2010

        doc = aw.Document(MY_DIR + "Document.docx", load_options)

        doc.save(ARTIFACTS_DIR + "WorkingWithLoadOptions.set_ms_word_version.docx")
        #ExEnd:SetMSWordVersion

    def test_use_temp_folder(self):

        #ExStart:UseTempFolder
        load_options = aw.loading.LoadOptions()
        load_options.temp_folder = ARTIFACTS_DIR

        doc = aw.Document(MY_DIR + "Document.docx", load_options)
        #ExEnd:UseTempFolder

    def test_load_with_encoding(self):

        #ExStart:LoadWithEncoding
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

        #ExStart:LoadCHM
        load_options = aw.loading.LoadOptions()
        load_options.encoding = "windows-1251"

        doc = aw.Document(MY_DIR + "HTML help.chm", load_options)
        #ExEnd:LoadCHM
