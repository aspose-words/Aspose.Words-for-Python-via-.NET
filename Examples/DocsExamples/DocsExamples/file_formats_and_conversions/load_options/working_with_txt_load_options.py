import io

import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class WorkingWithTxtLoadOptions(DocsExamplesBase):

    def test_detect_numbering_with_whitespaces(self):

        #ExStart:DetectNumberingWithWhitespaces
        # Create a plaintext document in the form of a string with parts that may be interpreted as lists.
        # Upon loading, the first three lists will always be detected by Aspose.words,
        # and List objects will be created for them after loading.
        text_doc = """
            Full stop delimiters:
            1. First list item 1
            2. First list item 2
            3. First list item 3

            Right bracket delimiters:
            1) Second list item 1
            2) Second list item 2
            3) Second list item 3

            Bullet delimiters:
            • Third list item 1
            • Third list item 2
            • Third list item 3

            Whitespace delimiters:
            1 Fourth list item 1
            2 Fourth list item 2
            3 Fourth list item 3"""

        # The fourth list, with whitespace inbetween the list number and list item contents,
        # will only be detected as a list if "DetectNumberingWithWhitespaces" in a LoadOptions object is set to true,
        # to avoid paragraphs that start with numbers being mistakenly detected as lists.
        load_options = aw.loading.TxtLoadOptions()
        load_options.detect_numbering_with_whitespaces = True

        # Load the document while applying LoadOptions as a parameter and verify the result.
        doc = aw.Document(io.BytesIO(text_doc.encode("utf-8")), load_options)

        doc.save(ARTIFACTS_DIR + "WorkingWithTxtLoadOptions.detect_numbering_with_whitespaces.docx")
        #ExEnd:DetectNumberingWithWhitespaces

    def test_handle_spaces_options(self):

        #ExStart:HandleSpacesOptions
        text_doc = "      Line 1 \n    Line 2   \n Line 3       "

        load_options = aw.loading.TxtLoadOptions()
        load_options.leading_spaces_options = aw.loading.TxtLeadingSpacesOptions.TRIM
        load_options.trailing_spaces_options = aw.loading.TxtTrailingSpacesOptions.TRIM

        doc = aw.Document(io.BytesIO(text_doc.encode("utf-8")), load_options)

        doc.save(ARTIFACTS_DIR + "WorkingWithTxtLoadOptions.handle_spaces_options.docx")
        #ExEnd:HandleSpacesOptions

    def test_document_text_direction(self):

        #ExStart:DocumentTextDirection
        load_options = aw.loading.TxtLoadOptions()
        load_options.document_direction = aw.loading.DocumentDirection.AUTO

        doc = aw.Document(MY_DIR + "Hebrew text.txt", load_options)

        paragraph = doc.first_section.body.first_paragraph
        print(paragraph.paragraph_format.bidi)

        doc.save(ARTIFACTS_DIR + "WorkingWithTxtLoadOptions.document_text_direction.docx")
        #ExEnd:DocumentTextDirection
