import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR


class HelloWorld(DocsExamplesBase):

    def test_simple_hello_world(self):

        #ExStart:HelloWorld
        #GistId:ff869587c650d2a27785d5cd598ae0b4
        doc_a = aw.Document()
        builder = aw.DocumentBuilder(doc_a)

        # Insert text to the document start.
        builder.move_to_document_start()
        builder.write("First Hello World paragraph")

        doc_b = aw.Document(MY_DIR + "Document.docx")
        # Add document B to the and of document A, preserving document B formatting.
        doc_a.appendDocument(doc_b, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        doc_a.save(ARTIFACTS_DIR + "HelloWorld.SimpleHelloWorld.pdf")
        #ExEnd:HelloWorld
