# -*- coding: utf-8 -*-

# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################

import aspose.words as aw
import os
import unittest
from api_example_base import ApiExampleBase, MY_DIR


class ExAI(ApiExampleBase):
    @unittest.skip("This test should be run manually to manage API requests amount")
    def test_ai_summarize(self):
        #ExStart: AiSummarize
        #ExFor:GoogleAiModel
        #ExFor:OpenAiModel
        #ExFor:IAiModelText
        #ExFor:IAiModelText.summarize(Document, SummarizeOptions)
        #ExFor:IAiModelText.summarize(Document[], SummarizeOptions)
        #ExFor:SummarizeOptions
        #ExFor:SummarizeOptions.summary_length
        #ExFor:SummaryLength
        #ExFor:AiModel
        #ExFor:AiModel.create(AiModelType)
        #ExFor:AiModel.with_api_key(String)
        #ExFor:AiModelType
        #ExSummary: Shows how to summarize text using OpenAI and Google models.
        first_doc = aw.Document(MyDir + "Big document.docx")
        second_doc = aw.Document(MyDir + "Document.docx")
        api_key = os.getenv("API_KEY")
        # Use OpenAI or Google generative language models.
        model = aw.ai.AiModel.create(aw.ai.AiModelType.GPT_4O_MINI).with_api_key(api_key).as_open_ai_model()
        options = aw.ai.SummarizeOptions()
        options.summary_length = aw.ai.SummaryLength.SHORT
        one_document_summary = model.summarize(first_doc, options)
        oneDocumentSummary.save(ArtifactsDir + "AI.AiSummarize.One.docx")
        options.summary_length = aw.ai.SummaryLength.LONG
        multi_document_summary = model.summarize([first_doc, second_doc], options)
        multiDocumentSummary.save(ArtifactsDir + "AI.AiSummarize.Multi.docx")
        #ExEnd:AiSummarize
