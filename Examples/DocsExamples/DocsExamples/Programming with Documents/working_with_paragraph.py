import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithParagraph(docs_base.DocsExamplesBase):
    
    def test_count_lines_in_paragraph(self) :

        #ExStart:CountLinesInParagraph
        document = aw.Document(docs_base.my_dir + "Bibliography.docx")

        collector = aw.layout.LayoutCollector(document)
        it = aw.layout.LayoutEnumerator(document)

        for paragraph in document.get_child_nodes(aw.NodeType.PARAGRAPH, True) :
         
            paragraph = paragraph.as_paragraph()
            paraBreak = collector.get_entity(paragraph)

            stop = None
            prevItem = paragraph.previous_sibling
            if (prevItem != None) :
                prevBreak = collector.get_entity(prevItem)
                if (prevItem.node_type == aw.NodeType.PARAGRAPH) :
                    it.current = collector.get_entity(prevItem) # para break
                    it.move_parent()    # last line
                    stop = it.current
                elif (prevItem.node_type == aw.NodeType.TABLE) :
                    table = prevItem.as_table()
                    it.current = collector.get_entity(table.last_row.last_cell.last_paragraph) # cell break
                    it.move_parent()    # cell
                    it.move_parent()    # row
                    stop = it.current
                else :
                    raise RuntimeError()

            it.current = paraBreak
            it.move_parent()

            # We move from line to line in a paragraph.
            # When paragraph spans multiple pages the we will follow across them.
            count = 1
            while (it.current != stop) :
                if (not it.move_previous_logical()) :
                    break
                count += 1

            MAX_CHARS = 16
            paraText = paragraph.get_text()
            if (len(paraText) > MAX_CHARS) :
                paraText = f"{paraText.substring(0, MAX_CHARS)}..."

            print(f"Paragraph '{paraText}' has {count} line(-s).")
        #ExEnd:CountLinesInParagraph
        
    

if __name__ == '__main__':
    unittest.main()