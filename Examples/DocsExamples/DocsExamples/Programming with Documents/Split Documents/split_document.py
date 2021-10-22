import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class SplitDocument(docs_base.DocsExamplesBase):

        def test_by_headings_html(self) :

            #ExStart:SplitDocumentByHeadingsHtml
            doc = aw.Document(docs_base.my_dir + "Rendering.docx")

            options = aw.saving.HtmlSaveOptions()

            # Split a document into smaller parts, in this instance split by heading.
            options.document_split_criteria = aw.saving.DocumentSplitCriteria.HEADING_PARAGRAPH

            doc.save(docs_base.artifacts_dir + "SplitDocument.by_headings_html.html", options)
            #ExEnd:SplitDocumentByHeadingsHtml


        def test_by_sections_html(self) :

            doc = aw.Document(docs_base.my_dir + "Rendering.docx")

            #ExStart:SplitDocumentBySectionsHtml
            options = aw.saving.HtmlSaveOptions()
            options.document_split_criteria = aw.saving.DocumentSplitCriteria.SECTION_BREAK
            #ExEnd:SplitDocumentBySectionsHtml

            doc.save(docs_base.artifacts_dir + "SplitDocument.by_sections_html.html", options)


        def test_by_sections(self) :

            #ExStart:SplitDocumentBySections
            doc = aw.Document(docs_base.my_dir + "Big document.docx")

            for i in range(0, doc.sections.count) :

                # Split a document into smaller parts, in this instance, split by section.
                section = doc.sections[i].clone()

                new_doc = aw.Document()
                new_doc.sections.clear()

                new_section = new_doc.import_node(section, True).as_section()
                new_doc.sections.add(new_section)

                # Save each section as a separate document.
                new_doc.save(docs_base.artifacts_dir + f"SplitDocument.by_sections_{i}.docx")

            #ExEnd:SplitDocumentBySections


        def test_page_by_page(self) :

            #ExStart:SplitDocumentPageByPage
            doc = aw.Document(docs_base.my_dir + "Big document.docx")

            page_count = doc.page_count

            for page in range(0, page_count) :

                # Save each page as a separate document.
                extracted_page = doc.extract_pages(page, 1)
                extracted_page.save(docs_base.artifacts_dir + f"SplitDocument.page_by_page_{page + 1}.docx")

            #ExEnd:SplitDocumentPageByPage

            self.merge_documents()


        #ExStart:MergeSplitDocuments
        @staticmethod
        def merge_documents() :

            # Find documents using for merge.
            document_paths = [f for f in os.listdir(docs_base.artifacts_dir) if (os.path.isfile(os.path.join(docs_base.artifacts_dir, f)) and f.find("SplitDocument.page_by_page_") >= 0)]

            source_document_path = os.path.join(docs_base.artifacts_dir, document_paths[0])

            # Open the first part of the resulting document.
            source_doc = aw.Document(source_document_path)

            # Create a new resulting document.
            merged_doc = aw.Document()
            merged_doc_builder = aw.DocumentBuilder(merged_doc)

            # Merge document parts one by one.
            for document_path in document_paths :

                document_path = os.path.join(docs_base.artifacts_dir, document_path)
                if (document_path == source_document_path) :
                    continue

                merged_doc_builder.move_to_document_end()
                merged_doc_builder.insert_document(source_doc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)
                source_doc = aw.Document(document_path)


            merged_doc.save(docs_base.artifacts_dir + "SplitDocument.merge_documents.docx")

        #ExEnd:MergeSplitDocuments

        def test_by_page_range(self) :

            #ExStart:SplitDocumentByPageRange
            doc = aw.Document(docs_base.my_dir + "Big document.docx")

            # Get part of the document.
            extracted_pages = doc.extract_pages(3, 6)
            extracted_pages.save(docs_base.artifacts_dir + "SplitDocument.by_page_range.docx")
            #ExEnd:SplitDocumentByPageRange


if __name__ == '__main__':
        unittest.main()
