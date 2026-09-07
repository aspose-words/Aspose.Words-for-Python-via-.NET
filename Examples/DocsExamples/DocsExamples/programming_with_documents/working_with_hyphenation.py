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

    #ExStart:CustomHyphenation
    #GistId:6cfc4dd3ee1b881f904b3ce31a3110f7
    def test_hyphenation_callback(self):

        try:
            # Register hyphenation callback.
            aw.Hyphenation.set_callback(self.CustomHyphenationCallback())

            document = aw.Document(MY_DIR + "German text.docx")
            document.save(ARTIFACTS_DIR + "WorkingWithHyphenation.hyphenation_callback.pdf")
        except RuntimeError as e:
            if str(e).startswith("Missing hyphenation dictionary"):
                print(str(e))
            else:
                raise
        finally:
            aw.Hyphenation.set_callback(None)

    class CustomHyphenationCallback(aw.IHyphenationCallback):

        def request_dictionary(self, language: str):
            if language == "en-US":
                dictionary_full_file_name = MY_DIR + "hyph_en_US.dic"
            elif language == "de-CH":
                dictionary_full_file_name = MY_DIR + "hyph_de_CH.dic"
            else:
                raise RuntimeError(f"Missing hyphenation dictionary for {language}.")

            # Register dictionary for requested language.
            aw.Hyphenation.register_dictionary(language, dictionary_full_file_name)
    #ExEnd:CustomHyphenation
