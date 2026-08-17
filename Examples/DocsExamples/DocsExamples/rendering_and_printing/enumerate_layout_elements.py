import aspose.words as aw

from docs_examples_base import DocsExamplesBase, MY_DIR

class EnumerateLayoutElements(DocsExamplesBase):

    def test_get_layout_elements(self):

        doc = aw.Document(MY_DIR + "Document layout.docx")

        # Enumerator which is used to "walk" the elements of a rendered document.
        layout_enumerator = aw.layout.LayoutEnumerator(doc)

        # Use the enumerator to write information about each layout element to the console.
        EnumerateLayoutElements.display_layout_elements(layout_enumerator, "")

    @staticmethod
    def display_layout_elements(layout_enumerator: aw.layout.LayoutEnumerator, padding: str):
        """Enumerates forward through each layout element in the document and prints out details of each element."""

        while True:
            EnumerateLayoutElements.display_entity_info(layout_enumerator, padding)

            if layout_enumerator.move_first_child():
                # Recurse into this child element.
                EnumerateLayoutElements.display_layout_elements(layout_enumerator, padding + " " * 4)
                layout_enumerator.move_parent()

            if not layout_enumerator.move_next():
                break

    @staticmethod
    def display_entity_info(layout_enumerator: aw.layout.LayoutEnumerator, padding: str):
        """Displays information about the current layout entity to the console."""

        info = f"{padding}{layout_enumerator.type} - {layout_enumerator.kind}"

        if layout_enumerator.type == aw.layout.LayoutEntityType.SPAN:
            info += f" - {layout_enumerator.text}"

        print(info)
