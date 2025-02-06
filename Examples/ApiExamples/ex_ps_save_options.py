# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
import aspose.words.saving
import aspose.words.settings
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExPsSaveOptions(ApiExampleBase):

    def test_use_book_fold_printing_settings(self):
        for render_text_as_book_fold in [False, True]:
            #ExStart
            #ExFor:PsSaveOptions
            #ExFor:PsSaveOptions.save_format
            #ExFor:PsSaveOptions.use_book_fold_printing_settings
            #ExSummary:Shows how to save a document to the Postscript format in the form of a book fold.
            doc = aw.Document(file_name=MY_DIR + 'Paragraphs.docx')
            # Create a "PsSaveOptions" object that we can pass to the document's "Save" method
            # to modify how that method converts the document to PostScript.
            # Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
            # in the output Postscript document in a way that helps us make a booklet out of it.
            # Set the "UseBookFoldPrintingSettings" property to "false" to save the document normally.
            save_options = aw.saving.PsSaveOptions()
            save_options.save_format = aw.SaveFormat.PS
            save_options.use_book_fold_printing_settings = render_text_as_book_fold
            # If we are rendering the document as a booklet, we must set the "MultiplePages"
            # properties of the page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
            for s in doc.sections:
                s = s.as_section()
                s.page_setup.multiple_pages = aw.settings.MultiplePagesType.BOOK_FOLD_PRINTING
            # Once we print this document on both sides of the pages, we can fold all the pages down the middle at once,
            # and the contents will line up in a way that creates a booklet.
            doc.save(file_name=ARTIFACTS_DIR + 'PsSaveOptions.UseBookFoldPrintingSettings.ps', save_options=save_options)
            #ExEnd