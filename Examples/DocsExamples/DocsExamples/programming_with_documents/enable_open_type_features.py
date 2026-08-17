import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class EnableOpenTypeFeatures(DocsExamplesBase):

    def test_open_type_features(self):
        #ExStart:OpenTypeFeatures
        #GistId:b5ab3801a7643f50529361fb177f61f5
        doc = aw.Document(file_name=MY_DIR + "OpenType text shaping.docx")
        doc.layout_options.enable_text_shaping = True
        doc.save(file_name=ARTIFACTS_DIR + 'EnableOpenTypeFeatures.EnableOpenTypeFeatures.pdf')
        #ExEnd:OpenTypeFeatures