import unittest
import os
import sys
import io

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

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

        stream = io.FileIO(MY_DIR + "hyph_de_CH.dic")
        aw.Hyphenation.register_dictionary("de-CH", stream)

        doc.save(ARTIFACTS_DIR + "WorkingWithHyphenation.load_hyphenation_dictionary_for_language.pdf")

        stream.close()
        #ExEnd:LoadHyphenationDictionaryForLanguage


if __name__ == '__main__':
    unittest.main()
