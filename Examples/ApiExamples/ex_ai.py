# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import os
import unittest
import aspose.words as aw
import aspose.words.ai
import system_helper
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExAI(ApiExampleBase):

    @unittest.skip('This test should be run manually to manage API requests amount')
    def test_ai_summarize(self):
        #ExStart:AiSummarize
        #ExFor:GoogleAiModel
        #ExFor:OpenAiModel
        #ExFor:OpenAiModel.with_organization(str)
        #ExFor:OpenAiModel.with_project(str)
        #ExFor:AiModel
        #ExFor:AiModel.summarize(Document,SummarizeOptions)
        #ExFor:AiModel.summarize(List[Document],SummarizeOptions)
        #ExFor:AiModel.create(AiModelType)
        #ExFor:AiModel.with_api_key(str)
        #ExFor:AiModelType
        #ExFor:SummarizeOptions
        #ExFor:SummarizeOptions.__init__
        #ExFor:SummarizeOptions.summary_length
        #ExFor:SummaryLength
        #ExSummary:Shows how to summarize text using OpenAI and Google models.
        first_doc = aw.Document(file_name=MY_DIR + 'Big document.docx')
        second_doc = aw.Document(file_name=MY_DIR + 'Document.docx')
        api_key = system_helper.environment.Environment.get_environment_variable('API_KEY')
        # Use OpenAI or Google generative language models.
        model = aw.ai.AiModel.create(aw.ai.AiModelType.GPT_4O_MINI).with_api_key(api_key).as_open_ai_model().with_organization('Organization').with_project('Project')
        options = aw.ai.SummarizeOptions()
        options.summary_length = aw.ai.SummaryLength.SHORT
        one_document_summary = model.summarize(source_document=first_doc, options=options)
        one_document_summary.save(file_name=ARTIFACTS_DIR + 'AI.AiSummarize.One.docx')
        options.summary_length = aw.ai.SummaryLength.LONG
        multi_document_summary = model.summarize(source_documents=[first_doc, second_doc], options=options)
        multi_document_summary.save(file_name=ARTIFACTS_DIR + 'AI.AiSummarize.Multi.docx')
        #ExEnd:AiSummarize

    @unittest.skip('This test should be run manually to manage API requests amount')
    def test_ai_translate(self):
        #ExStart:AiTranslate
        #ExFor:AiModel.translate(Document,AI.Language)
        #ExFor:AI.language
        #ExSummary:Shows how to translate text using Google models.
        doc = aw.Document(file_name=MY_DIR + 'Document.docx')
        api_key = system_helper.environment.Environment.get_environment_variable('API_KEY')
        # Use Google generative language models.
        model = aw.ai.AiModel.create(aw.ai.AiModelType.GEMINI_15_FLASH).with_api_key(api_key)
        translated_doc = model.translate(doc, aw.ai.Language.ARABIC)
        translated_doc.save(file_name=ARTIFACTS_DIR + 'AI.AiTranslate.docx')
        #ExEnd:AiTranslate

    @unittest.skip('This test should be run manually to manage API requests amount')
    def test_ai_grammar(self):
        #ExStart:AiGrammar
        #ExFor:AiModel.check_grammar(Document,CheckGrammarOptions)
        #ExFor:CheckGrammarOptions
        #ExSummary:Shows how to check the grammar of a document.
        doc = aw.Document(file_name=MY_DIR + 'Big document.docx')
        api_key = system_helper.environment.Environment.get_environment_variable('API_KEY')
        # Use OpenAI generative language models.
        model = aw.ai.AiModel.create(aw.ai.AiModelType.GPT_4O_MINI).with_api_key(api_key)
        grammar_options = aw.ai.CheckGrammarOptions()
        grammar_options.improve_stylistics = True
        proofed_doc = model.check_grammar(doc, grammar_options)
        proofed_doc.save(file_name=ARTIFACTS_DIR + 'AI.AiGrammar.docx')
        #ExEnd:AiGrammar