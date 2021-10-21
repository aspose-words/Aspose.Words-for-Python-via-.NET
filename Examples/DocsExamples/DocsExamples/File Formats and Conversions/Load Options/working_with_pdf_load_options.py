import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithPdfLoadOptions(docs_base.DocsExamplesBase):

    def test_load_encrypted_pdf(self) :

        #ExStart:LoadEncryptedPdf
        doc = aw.Document(docs_base.my_dir + "Pdf Document.pdf")

        saveOptions = aw.saving.PdfSaveOptions()
        saveOptions.encryption_details = aw.saving.PdfEncryptionDetails("Aspose", None, aw.saving.PdfEncryptionAlgorithm.RC4_40)


        doc.save(docs_base.artifacts_dir + "WorkingWithPdfLoadOptions.load_encrypted_pdf.pdf", saveOptions)

        loadOptions = aw.loading.PdfLoadOptions()
        loadOptions.password = "Aspose"
        loadOptions.load_format = aw.LoadFormat.PDF

        doc = aw.Document(docs_base.artifacts_dir + "WorkingWithPdfLoadOptions.load_encrypted_pdf.pdf", loadOptions)
        #ExEnd:LoadEncryptedPdf


    def test_load_page_range_of_pdf(self) :

        #ExStart:LoadPageRangeOfPdf
        loadOptions = aw.loading.PdfLoadOptions()
        loadOptions.page_index = 0;
        loadOptions.page_count = 1

        #ExStart:LoadPDF
        doc = aw.Document(docs_base.my_dir + "Pdf Document.pdf", loadOptions)

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfLoadOptions.load_page_range_of_pdf.pdf")
        #ExEnd:LoadPDF
        #ExEnd:LoadPageRangeOfPdf




if __name__ == '__main__':
    unittest.main()