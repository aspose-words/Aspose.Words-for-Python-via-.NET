import aspose.words as aw

import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR


class ExHarfBuzz(ApiExampleBase):


    def test_harf_buzz(self):
        #ExStart
        #ExFor:LayoutOptions.enable_text_shaping
        #ExSummary:Shows how to use enable_text_shaping property.
        doc =  aw.Document(file_name=MY_DIR + "TestMetricsKerning.docx")
        doc.layout_options.enable_text_shaping = True
        doc.save( file_name=ARTIFACTS_DIR + 'out_HarfBuzz.pdf')
        #ExEnd