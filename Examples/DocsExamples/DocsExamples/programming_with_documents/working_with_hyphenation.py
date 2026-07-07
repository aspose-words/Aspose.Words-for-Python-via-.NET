import io

import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class WorkingWithHyphenation(DocsExamplesBase):

    def test_hyphenate_words(self):

        #ExStart:HyphenateWords
        #GistId:6cfc4dd3ee1b881f904b3ce31a3110f7
        doc = aw.Document(MY_DIR + "German text.docx")

        aw.Hyphenation.register_dictionary("en-US", MY_DIR + "hyph_en_US.dic")
        aw.Hyphenation.register_dictionary("de-CH", MY_DIR + "hyph_de_CH.dic")

        doc.save(ARTIFACTS_DIR + "WorkingWithHyphenation.hyphenate_words.pdf")
        #ExEnd:HyphenateWords

    def test_load_hyphenation_dictionary(self):

        #ExStart:LoadHyphenationDictionary
        #GistId:6cfc4dd3ee1b881f904b3ce31a3110f7
        doc = aw.Document(MY_DIR + "German text.docx")

        with io.FileIO(MY_DIR + "hyph_de_CH.dic") as stream:
            aw.Hyphenation.register_dictionary("de-CH", stream)

        doc.save(ARTIFACTS_DIR + "WorkingWithHyphenation.load_hyphenation_dictionary.pdf")
        #ExEnd:LoadHyphenationDictionary
