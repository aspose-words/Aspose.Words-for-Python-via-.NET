import os
import unittest

import aspose.words as aw
import aspose.words.ai as aw_ai
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR


class WorkingWithAi(DocsExamplesBase):

    @unittest.skip("This test should be run manually to manage API requests amount")
    def test_ai_summarize(self):
        #ExStart:AiSummarize
        #GistId:1e379bedb2b759c1be24c64aad54d13d
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
        #GistId:ea14b3e44c0233eecd663f783a21c4f6
        doc = aw.Document(MY_DIR + "Document.docx")

        api_key = os.getenv("API_KEY")
        # Use Google generative language models.
        model = aw_ai.AiModel.create(aw_ai.AiModelType.GEMINI_15_FLASH).with_api_key(api_key).as_google_ai_model()

        translated_doc = model.translate(doc, aw_ai.Language.ARABIC)
        translated_doc.save(ARTIFACTS_DIR + "AI.AiTranslate.docx")
        #ExEnd:AiTranslate

    @unittest.skip("This test should be run manually to manage API requests amount")
    def test_ai_grammar(self):
        #ExStart:AiGrammar
        #GistId:98a646d19cd7708ed0cd3d97b993a053
        doc = aw.Document(MY_DIR + "Big document.docx")

        api_key = os.getenv("API_KEY")
        # Use OpenAI generative language models.
        model = aw_ai.AiModel.create(aw_ai.AiModelType.GPT_4O_MINI).with_api_key(api_key)

        grammar_options = aw_ai.CheckGrammarOptions()
        grammar_options.improve_stylistics = True

        proofed_doc = model.check_grammar(doc, grammar_options)
        proofed_doc.save(ARTIFACTS_DIR + "AI.AiGrammar.docx")
        #ExEnd:AiGrammar
