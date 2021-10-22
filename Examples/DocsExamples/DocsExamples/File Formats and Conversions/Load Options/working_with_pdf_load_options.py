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

    def test_load_encrypted_pdf(self):

        #ExStart:LoadEncryptedPdf
        doc = aw.Document(docs_base.my_dir + "Pdf Document.pdf")

        save_options = aw.saving.PdfSaveOptions()
        save_options.encryption_details = aw.saving.PdfEncryptionDetails("Aspose", None, aw.saving.PdfEncryptionAlgorithm.RC4_40)


        doc.save(docs_base.artifacts_dir + "WorkingWithPdfLoadOptions.load_encrypted_pdf.pdf", save_options)

        load_options = aw.loading.PdfLoadOptions()
        load_options.password = "Aspose"
        load_options.load_format = aw.LoadFormat.PDF

        doc = aw.Document(docs_base.artifacts_dir + "WorkingWithPdfLoadOptions.load_encrypted_pdf.pdf", load_options)
        #ExEnd:LoadEncryptedPdf


    def test_load_page_range_of_pdf(self):

        #ExStart:LoadPageRangeOfPdf
        load_options = aw.loading.PdfLoadOptions()
        load_options.page_index = 0;
        load_options.page_count = 1

        #ExStart:LoadPDF
        doc = aw.Document(docs_base.my_dir + "Pdf Document.pdf", load_options)

        doc.save(docs_base.artifacts_dir + "WorkingWithPdfLoadOptions.load_page_range_of_pdf.pdf")
        #ExEnd:LoadPDF
        #ExEnd:LoadPageRangeOfPdf


if __name__ == '__main__':
    unittest.main()
