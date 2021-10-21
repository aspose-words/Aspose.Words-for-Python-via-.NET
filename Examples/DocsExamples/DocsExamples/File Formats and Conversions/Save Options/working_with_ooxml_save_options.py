import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithOoxmlSaveOptions(docs_base.DocsExamplesBase):

    def test_encrypt_docx_with_password(self) :

        #ExStart:EncryptDocxWithPassword
        doc = aw.Document(docs_base.my_dir + "Document.docx")

        save_options = aw.saving.OoxmlSaveOptions()
        save_options.password = "password"

        doc.save(docs_base.artifacts_dir + "WorkingWithOoxmlSaveOptions.encrypt_docx_with_password.docx", save_options)
        #ExEnd:EncryptDocxWithPassword


    def test_ooxml_compliance_iso_29500_2008_strict(self) :

        #ExStart:OoxmlComplianceIso29500_2008_Strict
        doc = aw.Document(docs_base.my_dir + "Document.docx")

        doc.compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2016)

        save_options = aw.saving.OoxmlSaveOptions()
        save_options.compliance = aw.saving.OoxmlCompliance.ISO29500_2008_STRICT

        doc.save(docs_base.artifacts_dir + "WorkingWithOoxmlSaveOptions.ooxml_compliance_iso_29500_2008_strict.docx", save_options)
        #ExEnd:OoxmlComplianceIso29500_2008_Strict


    def test_update_last_saved_time_property(self) :

        #ExStart:UpdateLastSavedTimeProperty
        doc = aw.Document(docs_base.my_dir + "Document.docx")

        save_options = aw.saving.OoxmlSaveOptions()
        save_options.update_last_saved_time_property = True

        doc.save(docs_base.artifacts_dir + "WorkingWithOoxmlSaveOptions.update_last_saved_time_property.docx", save_options)
        #ExEnd:UpdateLastSavedTimeProperty


    def test_keep_legacy_control_chars(self) :

        #ExStart:KeepLegacyControlChars
        doc = aw.Document(docs_base.my_dir + "Legacy control character.doc")

        save_options = aw.saving.OoxmlSaveOptions(aw.SaveFormat.FLAT_OPC)
        save_options.keep_legacy_control_chars = True

        doc.save(docs_base.artifacts_dir + "WorkingWithOoxmlSaveOptions.keep_legacy_control_chars.docx", save_options)
        #ExEnd:KeepLegacyControlChars


    def test_set_compression_level(self) :

        #ExStart:SetCompressionLevel
        doc = aw.Document(docs_base.my_dir + "Document.docx")

        save_options = aw.saving.OoxmlSaveOptions()
        save_options.compression_level = aw.saving.CompressionLevel.SUPER_FAST

        doc.save(docs_base.artifacts_dir + "WorkingWithOoxmlSaveOptions.set_compression_level.docx", save_options)
        #ExEnd:SetCompressionLevel



if __name__ == '__main__':
    unittest.main()