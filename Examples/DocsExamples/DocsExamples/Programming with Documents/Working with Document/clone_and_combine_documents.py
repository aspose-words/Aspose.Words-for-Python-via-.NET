import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class CloneAndCombineDocuments(docs_base.DocsExamplesBase):

    def test_cloning_document(self) :
        
        #ExStart:CloningDocument
        doc = aw.Document(docs_base.my_dir + "Document.docx")

        clone = doc.clone().as_document()
        clone.save(docs_base.artifacts_dir + "CloneAndCombineDocuments.cloning_document.docx")
        #ExEnd:CloningDocument
 

    def test_insert_document_at_bookmark(self) :
        
        #ExStart:InsertDocumentAtBookmark         
        mainDoc = aw.Document(docs_base.my_dir + "Document insertion 1.docx")
        subDoc = aw.Document(docs_base.my_dir + "Document insertion 2.docx")

        bookmark = mainDoc.range.bookmarks.get_by_name("insertionPlace")
        self.insert_document(bookmark.bookmark_start.parent_node, subDoc)
            
        mainDoc.save(docs_base.artifacts_dir + "CloneAndCombineDocuments.insert_document_at_bookmark.docx")
        #ExEnd:InsertDocumentAtBookmark
        

    # <summary>
    # Inserts content of the external document after the specified node.
    # Section breaks and section formatting of the inserted document are ignored.
    # </summary>
    # <param name="insertionDestination">Node in the destination document after which the content
    # Should be inserted. This node should be a block level node (paragraph or table).</param>
    # <param name="docToInsert">The document to insert.</param>
    #ExStart:InsertDocument
    @staticmethod
    def insert_document(insertionDestination : aw.Node, docToInsert : aw.Document) :
        
        if (insertionDestination.node_type == aw.NodeType.PARAGRAPH or insertionDestination.node_type == awNodeType.TABLE) :
            
            destinationParent = insertionDestination.parent_node

            importer = aw.NodeImporter(docToInsert, insertionDestination.document, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

            # Loop through all block-level nodes in the section's body,
            # then clone and insert every node that is not the last empty paragraph of a section.
            for srcSection in docToInsert.sections :
                for srcNode in srcSection.as_section().body.child_nodes :
                    if (srcNode.node_type == aw.NodeType.PARAGRAPH) :
                    
                        para = srcNode.as_paragraph()
                        if (para.is_end_of_section and not para.has_child_nodes) :
                            continue

                    newNode = importer.import_node(srcNode, True)

                    destinationParent.insert_after(newNode, insertionDestination)
                    insertionDestination = newNode
        else :
            
            raise ValueError("The destination node should be either a paragraph or table.")
            
        
    #ExEnd:InsertDocument

    #ExStart:InsertDocumentWithSectionFormatting
    # <summary>
    # Inserts content of the external document after the specified node.
    # </summary>
    # <param name="insertAfterNode">Node in the destination document after which the content
    # Should be inserted. This node should be a block level node (paragraph or table).</param>
    # <param name="srcDoc">The document to insert.</param>
    @staticmethod
    def InsertDocumentWithSectionFormatting(insertAfterNode : aw.Node, srcDoc : aw.Document) :
        
        if (insertAfterNode.node_type != aw.NodeType.PARAGRAPH and
            insertAfterNode.node_type != aw.NodeType.TABLE) :
            raise ValueError("The destination node should be either a paragraph or table.")

        dstDoc = insertAfterNode.document.as_document()
        # To retain section formatting, split the current section into two at the marker node and then import the content
        # from srcDoc as whole sections. The section of the node to which the insert marker node belongs.
        currentSection = insertAfterNode.get_ancestor(aw.NodeType.SECTION).as_section()

        # Don't clone the content inside the section, we just want the properties of the section retained.
        cloneSection = currentSection.clone(False).as_section()

        # However, make sure the clone section has a body but no empty first paragraph.
        cloneSection.ensure_minimum()
        cloneSection.body.first_paragraph.remove()

        insertAfterNode.document.insert_after(cloneSection, currentSection)

        # Append all nodes after the marker node to the new section. This will split the content at the section level at.
        # The marker so the sections from the other document can be inserted directly.
        currentNode = insertAfterNode.next_sibling
        while (currentNode != None) :
            
            nextNode = currentNode.next_sibling
            cloneSection.body.append_child(currentNode)
            currentNode = nextNode
            
        # This object will be translating styles and lists during the import.
        importer = aw.NodeImporter(srcDoc, dstDoc, aw.ImportFormatMode.USE_DESTINATION_STYLES)

        for srcSection in srcDoc.sections :
            
            newNode = importer.import_node(srcSection, True)

            dstDoc.insert_after(newNode, currentSection)
            currentSection = newNode.as_section()
            
        
    #ExEnd:InsertDocumentWithSectionFormatting
    

if __name__ == '__main__':
    unittest.main()