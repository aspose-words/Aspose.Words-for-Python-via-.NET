from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

import aspose.words as aw

class WorkingWithRtfSaveOptions(DocsExamplesBase):

    def test_saving_images_as_wmf(self):

        #ExStart:SavingImagesAsWmf
        doc = aw.Document(MY_DIR + "Document.docx")

        save_options = aw.saving.RtfSaveOptions()
        save_options.save_images_as_wmf = True

        doc.save(ARTIFACTS_DIR + "WorkingWithRtfSaveOptions.saving_images_as_wmf.rtf", save_options)
        #ExEnd:SavingImagesAsWmf
