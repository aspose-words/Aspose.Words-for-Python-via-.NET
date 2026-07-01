import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class WorkingWithTableOfContent(DocsExamplesBase):

    def test_change_style_of_toc_level(self):

        #ExStart:ChangeStyleOfTocLevel
        #GistId:db118a3e1559b9c88355356df9d7ea10
        doc = aw.Document()
        # Retrieve the style used for the first level of the TOC and change the formatting of the style.
        doc.styles.get_by_style_identifier(aw.StyleIdentifier.TOC1).font.bold = True
        #ExEnd:ChangeStyleOfTocLevel

    def test_change_toc_tab_stops(self):

        #ExStart:ChangeTocTabStops
        #GistId:db118a3e1559b9c88355356df9d7ea10
        doc = aw.Document(MY_DIR + "Table of contents.docx")

        for para in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True):
            para = para.as_paragraph()
            # Check if this paragraph is formatted using the TOC result based styles.
            # This is any style between TOC and TOC9.
            if (para.paragraph_format.style.style_identifier >= aw.StyleIdentifier.TOC1 and
                para.paragraph_format.style.style_identifier <= aw.StyleIdentifier.TOC9):

                # Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
                tab = para.paragraph_format.tab_stops[0]

                # Remove the old tab from the collection.
                para.paragraph_format.tab_stops.remove_by_position(tab.position)

                # Insert a new tab using the same properties but at a modified position.
                # We could also change the separators used (dots) by passing a different Leader type.
                para.paragraph_format.tab_stops.add(tab.position - 50, tab.alignment, tab.leader)

        doc.save(ARTIFACTS_DIR + "WorkingWithTableOfContent.change_toc_tab_stops.docx")
        #ExEnd:ChangeTocTabStops

    def test_extract_toc(self):

        #ExStart:ExtractToc
        #GistId:db118a3e1559b9c88355356df9d7ea10
        doc = aw.Document(MY_DIR + "Table of contents.docx")

        for field in doc.range.fields:
            if field.type == aw.fields.FieldType.FIELD_HYPERLINK:
                hyperlink = field.as_field_hyperlink()
                if hyperlink.sub_address != None and hyperlink.sub_address.startswith("_Toc"):
                    toc_item = field.start.get_ancestor(aw.NodeType.PARAGRAPH).as_paragraph()
                    print(toc_item.to_string(aw.SaveFormat.TEXT).strip())
                    print("------------------")
                    if toc_item != None:
                        bm = doc.range.bookmarks.get_by_name(hyperlink.sub_address)
                        pointer = bm.bookmark_start.get_ancestor(aw.NodeType.PARAGRAPH).as_paragraph()
                        print(pointer.to_string(aw.SaveFormat.TEXT))
        #ExEnd:ExtractToc
