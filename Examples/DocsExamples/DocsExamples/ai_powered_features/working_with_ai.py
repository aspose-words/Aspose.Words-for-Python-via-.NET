import os
import unittest

import aspose.words as aw
import aspose.words.ai as aw_ai
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR


class WorkingWithAi(DocsExamplesBase):

    @unittest.skip("This test should be run manually to manage API requests amount")
    def test_ai_summarize(self):
        #ExStart:AiSummarize
        #GistId:c28e02ae4b77f1a92bc21aa9d79b5adc
        first_doc = aw.Document(MY_DIR + "Big document.docx")
        second_doc = aw.Document(MY_DIR + "Document.docx")

        api_key = os.getenv("API_KEY")
        # Use OpenAI or Google generative language models.
        model = aw_ai.AiModel.create(aw_ai.AiModelType.GPT_4O_MINI).with_api_key(api_key).as_open_ai_model().with_organization("Organization").with_project("Project")

        options = aw_ai.SummarizeOptions()

        options.summary_length = aw_ai.SummaryLength.SHORT
        one_document_summary = model.summarize(first_doc, options)
        one_document_summary.save(ARTIFACTS_DIR + "AI.AiSummarize.One.docx")

        options.summary_length = aw_ai.SummaryLength.LONG
        multi_document_summary = model.summarize([first_doc, second_doc], options)
        multi_document_summary.save(ARTIFACTS_DIR + "AI.AiSummarize.Multi.docx")
        #ExEnd:AiSummarize

    @unittest.skip("This test should be run manually to manage API requests amount")
    def test_ai_translate(self):
        #ExStart:AiTranslate
        #GistId:5e7ff3bb4165ea4255778ca6c65f3c51
        doc = aw.Document(MY_DIR + "Document.docx")

        api_key = os.getenv("API_KEY")
        # Use Google generative language models.
        model = aw_ai.AiModel.create(aw_ai.AiModelType.GEMINI_15_FLASH).with_api_key(api_key).as_google_ai_model()

        translated_doc = model.translate(doc, aw_ai.Language.ARABIC)
        translated_doc.save(ARTIFACTS_DIR + "AI.ai_translate.docx")
        #ExEnd:AiTranslate

    @unittest.skip("This test should be run manually to manage API requests amount")
    def test_ai_grammar(self):
        #ExStart:AiGrammar
        #GistId:e3ad1d3366734b5a6de2bb334702095d
        doc = aw.Document(MY_DIR + "Big document.docx")

        api_key = os.getenv("API_KEY")
        # Use OpenAI generative language models.
        model = aw_ai.AiModel.create(aw_ai.AiModelType.GPT_4O_MINI).with_api_key(api_key)

        grammar_options = aw_ai.CheckGrammarOptions()
        grammar_options.improve_stylistics = True

        proofed_doc = model.check_grammar(doc, grammar_options)
        proofed_doc.save(ARTIFACTS_DIR + "AI.ai_grammar.docx")
        #ExEnd:AiGrammar
