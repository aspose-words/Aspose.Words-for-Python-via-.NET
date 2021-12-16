import unittest
import io

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

class ExPsSaveOptions(ApiExampleBase):

    def test_use_book_fold_printing_settings(self):

        for render_text_as_book_fold in (False, True):
            with self.subTest(render_text_as_book_fold=render_text_as_book_fold):
                #ExStart
                #ExFor:PsSaveOptions
                #ExFor:PsSaveOptions.save_format
                #ExFor:PsSaveOptions.use_book_fold_printing_settings
                #ExSummary:Shows how to save a document to the Postscript format in the form of a book fold.
                doc = aw.Document(MY_DIR + "Paragraphs.docx")

                # Create a "PsSaveOptions" object that we can pass to the document's "save" method
                # to modify how that method converts the document to PostScript.
                # Set the "use_book_fold_printing_settings" property to "True" to arrange the contents
                # in the output Postscript document in a way that helps us make a booklet out of it.
                # Set the "use_book_fold_printing_settings" property to "False" to save the document normally.
                save_options = aw.saving.PsSaveOptions()
                save_options.save_format = aw.SaveFormat.PS
                save_options.use_book_fold_printing_settings = render_text_as_book_fold

                # If we are rendering the document as a booklet, we must set the "multiple_pages"
                # properties of the page setup objects of all sections to "MultiplePagesType.BOOK_FOLD_PRINTING".
                for s in doc.sections:
                    s = s.as_section()
                    s.page_setup.multiple_pages = aw.settings.MultiplePagesType.BOOK_FOLD_PRINTING

                # Once we print this document on both sides of the pages, we can fold all the pages down the middle at once,
                # and the contents will line up in a way that creates a booklet.
                doc.save(ARTIFACTS_DIR + "PsSaveOptions.use_book_fold_printing_settings.ps", save_options)
                #ExEnd
