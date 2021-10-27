from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

import aspose.words as aw

class WorkingWithOoxmlSaveOptions(DocsExamplesBase):

    def test_encrypt_docx_with_password(self):

        #ExStart:EncryptDocxWithPassword
        doc = aw.Document(MY_DIR + "Document.docx")

        save_options = aw.saving.OoxmlSaveOptions()
        save_options.password = "password"

        doc.save(ARTIFACTS_DIR + "WorkingWithOoxmlSaveOptions.encrypt_docx_with_password.docx", save_options)
        #ExEnd:EncryptDocxWithPassword

    def test_ooxml_compliance_iso_29500_2008_strict(self):

        #ExStart:OoxmlComplianceIso29500_2008_Strict
        doc = aw.Document(MY_DIR + "Document.docx")

        doc.compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2016)

        save_options = aw.saving.OoxmlSaveOptions()
        save_options.compliance = aw.saving.OoxmlCompliance.ISO29500_2008_STRICT

        doc.save(ARTIFACTS_DIR + "WorkingWithOoxmlSaveOptions.ooxml_compliance_iso_29500_2008_strict.docx", save_options)
        #ExEnd:OoxmlComplianceIso29500_2008_Strict

    def test_update_last_saved_time_property(self):

        #ExStart:UpdateLastSavedTimeProperty
        doc = aw.Document(MY_DIR + "Document.docx")

        save_options = aw.saving.OoxmlSaveOptions()
        save_options.update_last_saved_time_property = True

        doc.save(ARTIFACTS_DIR + "WorkingWithOoxmlSaveOptions.update_last_saved_time_property.docx", save_options)
        #ExEnd:UpdateLastSavedTimeProperty

    def test_keep_legacy_control_chars(self):

        #ExStart:KeepLegacyControlChars
        doc = aw.Document(MY_DIR + "Legacy control character.doc")

        save_options = aw.saving.OoxmlSaveOptions(aw.SaveFormat.FLAT_OPC)
        save_options.keep_legacy_control_chars = True

        doc.save(ARTIFACTS_DIR + "WorkingWithOoxmlSaveOptions.keep_legacy_control_chars.docx", save_options)
        #ExEnd:KeepLegacyControlChars

    def test_set_compression_level(self):

        #ExStart:SetCompressionLevel
        doc = aw.Document(MY_DIR + "Document.docx")

        save_options = aw.saving.OoxmlSaveOptions()
        save_options.compression_level = aw.saving.CompressionLevel.SUPER_FAST

        doc.save(ARTIFACTS_DIR + "WorkingWithOoxmlSaveOptions.set_compression_level.docx", save_options)
        #ExEnd:SetCompressionLevel
