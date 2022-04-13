import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class WorkingWithPdfLoadOptions(DocsExamplesBase):

    def test_load_encrypted_pdf(self):

        #ExStart:LoadEncryptedPdf
        doc = aw.Document(MY_DIR + "Pdf Document.pdf")

        save_options = aw.saving.PdfSaveOptions()
        save_options.encryption_details = aw.saving.PdfEncryptionDetails("Aspose", None)

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfLoadOptions.load_encrypted_pdf.pdf", save_options)

        load_options = aw.loading.PdfLoadOptions()
        load_options.password = "Aspose"
        load_options.load_format = aw.LoadFormat.PDF

        doc = aw.Document(ARTIFACTS_DIR + "WorkingWithPdfLoadOptions.load_encrypted_pdf.pdf", load_options)
        #ExEnd:LoadEncryptedPdf

    def test_load_page_range_of_pdf(self):

        #ExStart:LoadPageRangeOfPdf
        load_options = aw.loading.PdfLoadOptions()
        load_options.page_index = 0
        load_options.page_count = 1

        #ExStart:LoadPDF
        doc = aw.Document(MY_DIR + "Pdf Document.pdf", load_options)

        doc.save(ARTIFACTS_DIR + "WorkingWithPdfLoadOptions.load_page_range_of_pdf.pdf")
        #ExEnd:LoadPDF
        #ExEnd:LoadPageRangeOfPdf
