import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithTableOfContent(docs_base.DocsExamplesBase):

    def test_change_style_of_toc_level(self) :

        #ExStart:ChangeStyleOfTOCLevel
        doc = aw.Document()
        # Retrieve the style used for the first level of the TOC and change the formatting of the style.
        doc.styles.get_by_style_identifier(aw.StyleIdentifier.TOC1).font.bold = True
        #ExEnd:ChangeStyleOfTOCLevel


    def test_change_toc_tab_stops(self) :

        #ExStart:ChangeTOCTabStops
        doc = aw.Document(docs_base.my_dir + "Table of contents.docx")

        for para in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True) :
            para = para.as_paragraph()
            # Check if this paragraph is formatted using the TOC result based styles.
            # This is any style between TOC and TOC9.
            if (para.paragraph_format.style.style_identifier >= aw.StyleIdentifier.TOC1 and
                para.paragraph_format.style.style_identifier <= aw.StyleIdentifier.TOC9) :

                # Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
                tab = para.paragraph_format.tab_stops[0]

                # Remove the old tab from the collection.
                para.paragraph_format.tab_stops.remove_by_position(tab.position)

                # Insert a new tab using the same properties but at a modified position.
                # We could also change the separators used (dots) by passing a different Leader type.
                para.paragraph_format.tab_stops.add(tab.position - 50, tab.alignment, tab.leader)


        doc.save(docs_base.artifacts_dir + "WorkingWithTableOfContent.change_toc_tab_stops.docx")
        #ExEnd:ChangeTOCTabStops


if __name__ == '__main__':
    unittest.main()
