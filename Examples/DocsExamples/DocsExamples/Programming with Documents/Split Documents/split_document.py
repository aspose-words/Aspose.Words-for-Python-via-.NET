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

                newDoc = aw.Document()
                newDoc.sections.clear()

                newSection = newDoc.import_node(section, True).as_section()
                newDoc.sections.add(newSection)

                # Save each section as a separate document.
                newDoc.save(docs_base.artifacts_dir + f"SplitDocument.by_sections_{i}.docx")

            #ExEnd:SplitDocumentBySections


        def test_page_by_page(self) :

            #ExStart:SplitDocumentPageByPage
            doc = aw.Document(docs_base.my_dir + "Big document.docx")

            pageCount = doc.page_count

            for page in range(0, pageCount) :

                # Save each page as a separate document.
                extractedPage = doc.extract_pages(page, 1)
                extractedPage.save(docs_base.artifacts_dir + f"SplitDocument.page_by_page_{page + 1}.docx")

            #ExEnd:SplitDocumentPageByPage

            self.merge_documents()


        #ExStart:MergeSplitDocuments
        @staticmethod
        def merge_documents() :

            # Find documents using for merge.
            documentPaths = [f for f in os.listdir(docs_base.artifacts_dir) if (os.path.isfile(os.path.join(docs_base.artifacts_dir, f)) and f.find("SplitDocument.page_by_page_") >= 0)]

            sourceDocumentPath = os.path.join(docs_base.artifacts_dir, documentPaths[0])

            # Open the first part of the resulting document.
            sourceDoc = aw.Document(sourceDocumentPath)

            # Create a new resulting document.
            mergedDoc = aw.Document()
            mergedDocBuilder = aw.DocumentBuilder(mergedDoc)

            # Merge document parts one by one.
            for documentPath in documentPaths :

                documentPath = os.path.join(docs_base.artifacts_dir, documentPath)
                if (documentPath == sourceDocumentPath) :
                    continue

                mergedDocBuilder.move_to_document_end()
                mergedDocBuilder.insert_document(sourceDoc, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)
                sourceDoc = aw.Document(documentPath)


            mergedDoc.save(docs_base.artifacts_dir + "SplitDocument.merge_documents.docx")

        #ExEnd:MergeSplitDocuments

        def test_by_page_range(self) :

            #ExStart:SplitDocumentByPageRange
            doc = aw.Document(docs_base.my_dir + "Big document.docx")

            # Get part of the document.
            extractedPages = doc.extract_pages(3, 6)
            extractedPages.save(docs_base.artifacts_dir + "SplitDocument.by_page_range.docx")
            #ExEnd:SplitDocumentByPageRange





if __name__ == '__main__':
        unittest.main()