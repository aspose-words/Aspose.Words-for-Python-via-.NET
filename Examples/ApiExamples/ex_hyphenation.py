# Copyright (c) 2001-2022 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import aspose.words as aw

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExHyphenation(ApiExampleBase):

    def test_dictionary(self):

        #ExStart
        #ExFor:Hyphenation.is_dictionary_registered(str)
        #ExFor:Hyphenation.register_dictionary(str,str)
        #ExFor:Hyphenation.unregister_dictionary(str)
        #ExSummary:Shows how to register a hyphenation dictionary.
        # A hyphenation dictionary contains a list of strings that define hyphenation rules for the dictionary's language.
        # When a document contains lines of text in which a word could be split up and continued on the next line,
        # hyphenation will look through the dictionary's list of strings for that word's substrings.
        # If the dictionary contains a substring, then hyphenation will split the word across two lines
        # by the substring and add a hyphen to the first half.
        # Register a dictionary file from the local file system to the "de-CH" locale.
        aw.Hyphenation.register_dictionary("de-CH", MY_DIR + "hyph_de_CH.dic")

        self.assertTrue(aw.Hyphenation.is_dictionary_registered("de-CH"))

        # Open a document containing text with a locale matching that of our dictionary,
        # and save it to a fixed-page save format. The text in that document will be hyphenated.
        doc = aw.Document(MY_DIR + "German text.docx")

        self.assertTrue(all(node for node in doc.first_section.body.first_paragraph.runs
                            if node.as_run().font.locale_id == 2055))

        doc.save(ARTIFACTS_DIR + "Hyphenation.dictionary.registered.pdf")

        # Re-load the document after un-registering the dictionary,
        # and save it to another PDF, which will not have hyphenated text.
        aw.Hyphenation.unregister_dictionary("de-CH")

        self.assertFalse(aw.Hyphenation.is_dictionary_registered("de-CH"))

        doc = aw.Document(MY_DIR + "German text.docx")
        doc.save(ARTIFACTS_DIR + "Hyphenation.dictionary.unregistered.pdf")
        #ExEnd

        #pdf_doc = aspose.pdf.Document(ARTIFACTS_DIR + "Hyphenation.dictionary.registered.pdf")
        #text_absorber = aspose.pdf.text.TextAbsorber()
        #text_absorber.visit(pdf_doc)

        #self.assertIn(
        #    "La ob storen an deinen am sachen. Dop-\r\n" +
        #    "pelte  um  da  am  spateren  verlogen  ge-\r\n" +
        #    "kommen  achtzehn  blaulich.",
        #    text_absorber.text)

        #pdf_doc = aspose.pdf.Document(ARTIFACTS_DIR + "Hyphenation.dictionary.unregistered.pdf")
        #text_absorber = aspose.pdf.text.TextAbsorber()
        #text_absorber.visit(pdf_doc)

        #self.assertIn(
        #    "La  ob  storen  an  deinen  am  sachen. \r\n" +
        #    "Doppelte  um  da  am  spateren  verlogen \r\n" +
        #    "gekommen  achtzehn  blaulich.",
        #    text_absorber.text)

    ##ExStart
    ##ExFor:Hyphenation
    ##ExFor:Hyphenation.callback
    ##ExFor:Hyphenation.register_dictionary(str,BytesIO)
    ##ExFor:Hyphenation.register_dictionary(str,str)
    ##ExFor:Hyphenation.warning_callback
    ##ExFor:IHyphenationCallback
    ##ExFor:IHyphenationCallback.request_dictionary(str)
    ##ExSummary:Shows how to open and register a dictionary from a file.
    #def test_register_dictionary(self):

    #    # Set up a callback that tracks warnings that occur during hyphenation dictionary registration.
    #    warning_info_collection = aw.WarningInfoCollection()
    #    aw.Hyphenation.warning_callback = warning_info_collection

    #    # Register an English (US) hyphenation dictionary by stream.
    #    dictionary_stream = open(MY_DIR + "hyph_en_US.dic", "rb")
    #    aw.Hyphenation.register_dictionary("en-US", dictionary_stream)

    #    self.assertEqual(0, warning_info_collection.count)

    #    # Open a document with a locale that Microsoft Word may not hyphenate on an English machine, such as German.
    #    doc = aw.Document(MY_DIR + "German text.docx")

    #    # To hyphenate that document upon saving, we need a hyphenation dictionary for the "de-CH" language code.
    #    # This callback will handle the automatic request for that dictionary.
    #    aw.Hyphenation.callback = ExHyphenation.CustomHyphenationDictionaryRegister()

    #    # When we save the document, German hyphenation will take effect.
    #    doc.save(ARTIFACTS_DIR + "Hyphenation.register_dictionary.pdf")

    #    # This dictionary contains two identical patterns, which will trigger a warning.
    #    self.assertEqual(1, warning_info_collection.count)
    #    self.assertEqual(aw.WarningType.MINOR_FORMATTING_LOSS, warning_info_collection[0].warning_type)
    #    self.assertEqual(aw.WarningSource.LAYOUT, warning_info_collection[0].source)
    #    self.assertEqual("Hyphenation dictionary contains duplicate patterns. The only first found pattern will be used. " +
    #                    "Content can be wrapped differently.", warning_info_collection[0].description)

    #class CustomHyphenationDictionaryRegister(aw.IHyphenationCallback):
    #    """Associates ISO language codes with local system filenames for hyphenation dictionary files."""

    #    def __init__(self):
    #        self.hyphenation_dictionary_files = {
    #            "en-US": MY_DIR + "hyph_en_US.dic",
    #            "de-CH": MY_DIR + "hyph_de_CH.dic",
    #            }

    #    def request_dictionary(self, language: str):

    #        print("Hyphenation dictionary requested: " + language, end="")

    #        if aw.Hyphenation.is_dictionary_registered(language):
    #            print(", is already registered.")
    #            return

    #        if self.hyphenation_dictionary_files.contains_key(language):
    #            aw.Hyphenation.register_dictionary(language, self.hyphenation_dictionary_files[language])
    #            print(", successfully registered.")
    #            return

    #        print(", no respective dictionary file known by this Callback.")

    ##ExEnd
