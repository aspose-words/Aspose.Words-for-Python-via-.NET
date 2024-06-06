import aspose.words as aw
from DocsExamples.docs_examples_base import DocsExamplesBase


class VideoExamples(DocsExamplesBase):

    def test_video_example(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello World!")

        doc.save("video_examples.doc")
