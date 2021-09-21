import unittest
import os
import sys
import io

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithHyphenation(docs_base.DocsExamplesBase):
    
    def test_hyphenate_words_of_languages(self) :
        
        #ExStart:HyphenateWordsOfLanguages
        doc = aw.Document(docs_base.my_dir + "German text.docx")

        aw.Hyphenation.register_dictionary("en-US", docs_base.my_dir + "hyph_en_US.dic")
        aw.Hyphenation.register_dictionary("de-CH", docs_base.my_dir + "hyph_de_CH.dic")

        doc.save(docs_base.artifacts_dir + "WorkingWithHyphenation.hyphenate_words_of_languages.pdf")
        #ExEnd:HyphenateWordsOfLanguages
        

    def test_load_hyphenation_dictionary_for_language(self) :
        
        #ExStart:LoadHyphenationDictionaryForLanguage
        doc = aw.Document(docs_base.my_dir + "German text.docx")
            
        stream = io.FileIO(docs_base.my_dir + "hyph_de_CH.dic")
        aw.Hyphenation.register_dictionary("de-CH", stream)

        doc.save(docs_base.artifacts_dir + "WorkingWithHyphenation.load_hyphenation_dictionary_for_language.pdf")

        stream.close()
        #ExEnd:LoadHyphenationDictionaryForLanguage
        

    ##ExStart:CustomHyphenation
    #def test_hyphenation_callback()
        
    #    try
            
    #        # Register hyphenation callback.
    #        Hyphenation.callback = new CustomHyphenationCallback()

    #        Document document = new Document(docs_base.my_dir + "German text.docx")
    #        document.save(docs_base.artifacts_dir + "WorkingWithHyphenation.hyphenation_callback.pdf")
            
    #    catch (Exception e) when (e.message.starts_with("Missing hyphenation dictionary"))
            
    #        print(e.message)
            
    #    finally
            
    #        Hyphenation.callback = null
            
        

    #public class CustomHyphenationCallback : IHyphenationCallback
        
    #    public void RequestDictionary(string language)
            
    #        string dictionaryFolder = docs_base.my_dir
    #        string dictionaryFullFileName
    #        switch (language)
                
    #            case "en-US":
    #                dictionaryFullFileName = Path.combine(dictionaryFolder, "hyph_en_US.dic")
    #                break
    #            case "de-CH":
    #                dictionaryFullFileName = Path.combine(dictionaryFolder, "hyph_de_CH.dic")
    #                break
    #            default:
    #                throw new Exception($"Missing hyphenation dictionary for language.")
                
    #        # Register dictionary for requested language.
    #        Hyphenation.register_dictionary(language, dictionaryFullFileName)
            
        
    ##ExEnd:CustomHyphenation
    

if __name__ == '__main__':
    unittest.main()