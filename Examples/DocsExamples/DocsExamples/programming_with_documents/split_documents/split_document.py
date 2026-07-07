import os

import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class SplitDocument(DocsExamplesBase):

    def test_by_headings_html(self):

        #ExStart:SplitDocumentByHeadings
        #GistId:c5d01a71f65e540e5e77650b846b24cc
        doc = aw.Document(MY_DIR + "Rendering.docx")

        options = aw.saving.HtmlSaveOptions()

        # Split a document into smaller parts, in this instance split by heading.
        options.document_split_criteria = aw.saving.DocumentSplitCriteria.HEADING_PARAGRAPH

        doc.save(ARTIFACTS_DIR + "SplitDocument.by_headings_html.html", options)
        #ExEnd:SplitDocumentByHeadings

    def test_by_sections_html(self):

        doc = aw.Document(MY_DIR + "Rendering.docx")

        #ExStart:SplitDocumentBySectionsHtml
        #GistId:fec4c47291463c8d8d8a6684ffb23d64
        options = aw.saving.HtmlSaveOptions()
        options.document_split_criteria = aw.saving.DocumentSplitCriteria.SECTION_BREAK
        #ExEnd:SplitDocumentBySectionsHtml

        doc.save(ARTIFACTS_DIR + "SplitDocument.by_sections_html.html", options)

    def test_by_sections(self):

        #ExStart:SplitDocumentBySections
        #GistId:fec4c47291463c8d8d8a6684ffb23d64
        doc = aw.Document(MY_DIR + "Big document.docx")

        for i in range(doc.sections.count):
            # Split a document into smaller parts, in this instance, split by section.
            section = doc.sections[i].clone()

            new_doc = aw.Document()
            new_doc.sections.clear()

            new_section = new_doc.import_node(section, True).as_section()
            new_doc.sections.add(new_section)

            # Save each section as a separate document.
            new_doc.save(ARTIFACTS_DIR + "SplitDocument.by_sections_{}.docx".format(i))

        #ExEnd:SplitDocumentBySections

    def test_page_by_page(self):

        #ExStart:SplitDocumentPageByPage
        #GistId:fec4c47291463c8d8d8a6684ffb23d64
        doc = aw.Document(MY_DIR + "Big document.docx")

        page_count = doc.page_count

        for page in range(page_count):
            # Save each page as a separate document.
            extracted_page = doc.extract_pages(page, 1)
            extracted_page.save(ARTIFACTS_DIR + "SplitDocument.page_by_page_{}.docx".format(page + 1))

        #ExEnd:SplitDocumentPageByPage

        self.merge_documents()

    #ExStart:MergeSplitDocuments
    #GistId:fec4c47291463c8d8d8a6684ffb23d64
    @staticmethod
    def merge_documents():

        # Find documents using for merge.
        document_paths = [f for f in os.listdir(ARTIFACTS_DIR)
                          if (os.path.isfile(os.path.join(ARTIFACTS_DIR, f)) and f.startswith("SplitDocument.page_by_page_"))]

        source_document_path = os.path.join(ARTIFACTS_DIR, document_paths[0])

        # Open the first part of the resulting document.
        source_doc = aw.Document(source_document_path)

        # Create a new resulting document.
        merged_doc = aw.Document()
        merged_doc_builder = aw.DocumentBuilder(merged_doc)

        # Merge document parts one by one.
        for document_path in document_paths:
            document_path = os.path.join(ARTIFACTS_DIR, document_path)
            if document_path == source_document_path:
                continue

            merged_doc_builder.move_to_document_end()
            merged_doc_builder.insert_document(source_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)
            source_doc = aw.Document(document_path)

        merged_doc.save(ARTIFACTS_DIR + "SplitDocument.merge_documents.docx")

    #ExEnd:MergeSplitDocuments

    def test_by_page_range(self):

        #ExStart:SplitDocumentByPageRange
        #GistId:fec4c47291463c8d8d8a6684ffb23d64
        doc = aw.Document(MY_DIR + "Big document.docx")

        # Get part of the document.
        extracted_pages = doc.extract_pages(3, 6)
        extracted_pages.save(ARTIFACTS_DIR + "SplitDocument.by_page_range.docx")
        #ExEnd:SplitDocumentByPageRange
