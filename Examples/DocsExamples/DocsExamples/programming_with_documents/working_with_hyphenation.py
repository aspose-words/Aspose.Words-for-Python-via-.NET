import unittest
import os
import sys
import io

from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

import aspose.words as aw

class WorkingWithHyphenation(DocsExamplesBase):

    def test_hyphenate_words_of_languages(self):

        #ExStart:HyphenateWordsOfLanguages
        doc = aw.Document(MY_DIR + "German text.docx")

        aw.Hyphenation.register_dictionary("en-US", MY_DIR + "hyph_en_US.dic")
        aw.Hyphenation.register_dictionary("de-CH", MY_DIR + "hyph_de_CH.dic")

        doc.save(ARTIFACTS_DIR + "WorkingWithHyphenation.hyphenate_words_of_languages.pdf")
        #ExEnd:HyphenateWordsOfLanguages

    def test_load_hyphenation_dictionary_for_language(self):

        #ExStart:LoadHyphenationDictionaryForLanguage
        doc = aw.Document(MY_DIR + "German text.docx")

        with io.FileIO(MY_DIR + "hyph_de_CH.dic") as stream:
            aw.Hyphenation.register_dictionary("de-CH", stream)

        doc.save(ARTIFACTS_DIR + "WorkingWithHyphenation.load_hyphenation_dictionary_for_language.pdf")
        #ExEnd:LoadHyphenationDictionaryForLanguage
