# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import aspose.words as aw

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExLayout(ApiExampleBase):

    def test_layout_collector(self):

        #ExStart
        #ExFor:LayoutCollector
        #ExFor:LayoutCollector.__init__(Document)
        #ExFor:LayoutCollector.clear
        #ExFor:LayoutCollector.document
        #ExFor:LayoutCollector.get_end_page_index(Node)
        #ExFor:LayoutCollector.get_entity(Node)
        #ExFor:LayoutCollector.get_num_pages_spanned(Node)
        #ExFor:LayoutCollector.get_start_page_index(Node)
        #ExFor:LayoutEnumerator.current
        #ExSummary:Shows how to see the the ranges of pages that a node spans.
        doc = aw.Document()
        layout_collector = aw.layout.LayoutCollector(doc)

        # Call the "get_num_pages_spanned" method to count how many pages the content of our document spans.
        # Since the document is empty, that number of pages is currently zero.
        self.assertEqual(doc, layout_collector.document)
        self.assertEqual(0, layout_collector.get_num_pages_spanned(doc))

        # Populate the document with 5 pages of content.
        builder = aw.DocumentBuilder(doc)
        builder.write("Section 1")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.insert_break(aw.BreakType.SECTION_BREAK_EVEN_PAGE)
        builder.write("Section 2")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.insert_break(aw.BreakType.PAGE_BREAK)

        # Before the layout collector, we need to call the "update_page_layout" method to give us
        # an accurate figure for any layout-related metric, such as the page count.
        self.assertEqual(0, layout_collector.get_num_pages_spanned(doc))

        layout_collector.clear()
        doc.update_page_layout()

        self.assertEqual(5, layout_collector.get_num_pages_spanned(doc))

        # We can see the numbers of the start and end pages of any node and their overall page spans.
        nodes = doc.get_child_nodes(aw.NodeType.ANY, True)
        for node in nodes:
            print(f"->  NodeType.{node.node_type}: ")
            print(
                f"\tStarts on page {layout_collector.get_start_page_index(node)}, ends on page {layout_collector.get_end_page_index(node)}," +
                f" spanning {layout_collector.get_num_pages_spanned(node)} pages.")

        # We can iterate over the layout entities using a LayoutEnumerator.
        layout_enumerator = aw.layout.LayoutEnumerator(doc)

        self.assertEqual(aw.layout.LayoutEntityType.PAGE, layout_enumerator.type)

        # The LayoutEnumerator can traverse the collection of layout entities like a tree.
        # We can also apply it to any node's corresponding layout entity.
        layout_enumerator.current = layout_collector.get_entity(doc.get_child(aw.NodeType.PARAGRAPH, 1, True))

        self.assertEqual(aw.layout.LayoutEntityType.SPAN, layout_enumerator.type)
        self.assertEqual("¶", layout_enumerator.text)
        #ExEnd

    #ExStart
    #ExFor:LayoutEntityType
    #ExFor:LayoutEnumerator
    #ExFor:LayoutEnumerator.__init__(Document)
    #ExFor:LayoutEnumerator.document
    #ExFor:LayoutEnumerator.kind
    #ExFor:LayoutEnumerator.move_first_child
    #ExFor:LayoutEnumerator.move_last_child
    #ExFor:LayoutEnumerator.move_next
    #ExFor:LayoutEnumerator.move_next_logical
    #ExFor:LayoutEnumerator.move_parent(LayoutEntityType)
    #ExFor:LayoutEnumerator.move_previous
    #ExFor:LayoutEnumerator.move_previous_logical
    #ExFor:LayoutEnumerator.page_index
    #ExFor:LayoutEnumerator.rectangle
    #ExFor:LayoutEnumerator.reset
    #ExFor:LayoutEnumerator.text
    #ExFor:LayoutEnumerator.type
    #ExSummary:Shows ways of traversing a document's layout entities.
    def test_layout_enumerator(self):

        # Open a document that contains a variety of layout entities.
        # Layout entities are pages, cells, rows, lines, and other objects included in the LayoutEntityType enum.
        # Each layout entity has a rectangular space that it occupies in the document body.
        doc = aw.Document(MY_DIR + "Layout entities.docx")

        # Create an enumerator that can traverse these entities like a tree.
        layout_enumerator = aw.layout.LayoutEnumerator(doc)

        self.assertEqual(doc, layout_enumerator.document)

        layout_enumerator.move_parent(aw.layout.LayoutEntityType.PAGE)

        self.assertEqual(aw.layout.LayoutEntityType.PAGE, layout_enumerator.type)
        with self.assertRaises(Exception):
            print(layout_enumerator.text)

        # We can call this method to make sure that the enumerator will be at the first layout entity.
        layout_enumerator.reset()

        # There are two orders that determine how the layout enumerator continues traversing layout entities
        # when it encounters entities that span across multiple pages.
        # 1 -  In visual order:
        # When moving through an entity's children that span multiple pages,
        # page layout takes precedence, and we move to other child elements on this page and avoid the ones on the next.
        print("Traversing from first to last, elements between pages separated:")
        ExLayout.traverse_layout_forward(layout_enumerator, 1)

        # Our enumerator is now at the end of the collection. We can traverse the layout entities backwards to go back to the beginning.
        print("Traversing from last to first, elements between pages separated:")
        ExLayout.traverse_layout_backward(layout_enumerator, 1)

        # 2 -  In logical order:
        # When moving through an entity's children that span multiple pages,
        # the enumerator will move between pages to traverse all the child entities.
        print("Traversing from first to last, elements between pages mixed:")
        ExLayout.traverse_layout_forward_logical(layout_enumerator, 1)

        print("Traversing from last to first, elements between pages mixed:")
        ExLayout.traverse_layout_backward_logical(layout_enumerator, 1)

    @staticmethod
    def traverse_layout_forward(layout_enumerator: aw.layout.LayoutEnumerator, depth: int):
        """Enumerate through layout_enumerator's layout entity collection front-to-back,
        in a depth-first manner, and in the "Visual" order."""

        while True:
            ExLayout.print_current_entity(layout_enumerator, depth)

            if layout_enumerator.move_first_child():
                ExLayout.traverse_layout_forward(layout_enumerator, depth + 1)
                layout_enumerator.move_parent()

            if not layout_enumerator.move_next():
                break

    @staticmethod
    def traverse_layout_backward(layout_enumerator: aw.layout.LayoutEnumerator, depth: int):
        """Enumerate through layout_enumerator's layout entity collection back-to-front,
        in a depth-first manner, and in the "Visual" order."""

        while True:
            ExLayout.print_current_entity(layout_enumerator, depth)

            if layout_enumerator.move_last_child():
                ExLayout.traverse_layout_backward(layout_enumerator, depth + 1)
                layout_enumerator.move_parent()

            if not layout_enumerator.move_previous():
                break

    @staticmethod
    def traverse_layout_forward_logical(layout_enumerator: aw.layout.LayoutEnumerator, depth: int):
        """Enumerate through layout_enumerator's layout entity collection front-to-back,
        in a depth-first manner, and in the "Logical" order."""

        while True:
            ExLayout.print_current_entity(layout_enumerator, depth)

            if layout_enumerator.move_first_child():
                ExLayout.traverse_layout_forward_logical(layout_enumerator, depth + 1)
                layout_enumerator.move_parent()

            if not layout_enumerator.move_next_logical():
                break

    @staticmethod
    def traverse_layout_backward_logical(layout_enumerator: aw.layout.LayoutEnumerator, depth: int):
        """Enumerate through layout_enumerator's layout entity collection back-to-front,
        in a depth-first manner, and in the "Logical" order."""

        while True:
            ExLayout.print_current_entity(layout_enumerator, depth)

            if layout_enumerator.move_last_child():
                ExLayout.traverse_layout_backward_logical(layout_enumerator, depth + 1)
                layout_enumerator.move_parent()

            if not layout_enumerator.move_previous_logical():
                break

    @staticmethod
    def print_current_entity(layout_enumerator: aw.layout.LayoutEnumerator, indent: int):
        """Print information about layout_enumerator's current entity to the console, while indenting the text with tab characters
        based on its depth relative to the root node that we provided in the constructor LayoutEnumerator instance.
        The rectangle that we process at the end represents the area and location that the entity takes up in the document."""

        tabs = "\t" * indent

        if layout_enumerator.kind == "":
            print(f"{tabs}-> Entity type: {layout_enumerator.type}")
        else:
            print(f"{tabs}-> Entity type & kind: {layout_enumerator.type}, {layout_enumerator.kind}")

        # Only spans can contain text.
        if layout_enumerator.type == aw.layout.LayoutEntityType.SPAN:
            print(f"{tabs}   Span contents: \"{layout_enumerator.text}\"")

        le_rect = layout_enumerator.rectangle
        print(f"{tabs}   Rectangle dimensions {le_rect.width}x{le_rect.height}, X={le_rect.x} Y={le_rect.y}")
        print(f"{tabs}   Page {layout_enumerator.page_index}")

    #ExEnd

    ##ExStart
    ##ExFor:IPageLayoutCallback
    ##ExFor:IPageLayoutCallback.notify(PageLayoutCallbackArgs)
    ##ExFor:PageLayoutCallbackArgs.event
    ##ExFor:PageLayoutCallbackArgs.document
    ##ExFor:PageLayoutCallbackArgs.page_index
    ##ExFor:PageLayoutEvent
    ##ExSummary:Shows how to track layout changes with a layout callback.

    #def test_page_layout_callback(self):

    #    doc = aw.Document()
    #    doc.built_in_document_properties.title = "My Document"

    #    builder = aw.DocumentBuilder(doc)
    #    builder.writeln("Hello world!")

    #    doc.layout_options.callback = ExLayout.RenderPageLayoutCallback()
    #    doc.update_page_layout()

    #    doc.save(ARTIFACTS_DIR + "Layout.page_layout_callback.pdf")

    #class RenderPageLayoutCallback(aw.layout.IPageLayoutCallback):
    #    """Notifies us when we save the document to a fixed page format
    #    and renders a page that we perform a page reflow on to an image in the local file system."""

    #    def __init__(self):
    #        self.num = 0

    #    def notify(self, a: aw.layout.PageLayoutCallbackArgs):

    #        if a.event == aw.layout.PageLayoutEvent.PART_REFLOW_FINISHED:
    #            self._notify_part_finished(a)
    #        elif a.event == aw.layout.PageLayoutEvent.CONVERSION_FINISHED:
    #            self._notify_conversion_finished(a)

    #    def _notify_part_finished(self, a: aw.layout.PageLayoutCallbackArgs):

    #        print(f"Part at page {a.PageIndex + 1} reflow.")
    #        self._render_page(a, a.page_index)

    #    def _notify_conversion_finished(self, a: aw.layout.PageLayoutCallbackArgs):

    #        print(f"Document \"{a.document.built_in_document_properties.title}\" converted to page format.")

    #    def _render_page(self, a: aw.layout.PageLayoutCallbackArgs, page_index: int):

    #        save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
    #        save_options.page_set = aw.saving.PageSet(page_index)

    #        self.num += 1
    #        with open(ARTIFACTS_DIR + "PageLayoutCallback.page-{page_index + 1} {self.num}.png", "wb") as stream:
    #            a.document.save(stream, save_options)

    ##ExEnd

    def test_restart_page_numbering_in_continuous_section(self):

        #ExStart
        #ExFor:LayoutOptions.continuous_section_page_numbering_restart
        #ExFor:ContinuousSectionRestart
        #ExSummary:Shows how to control page numbering in a continuous section.
        doc = aw.Document(MY_DIR + "Continuous section page numbering.docx")

        # By default Aspose.Words behavior matches the Microsoft Word 2019.
        # If you need old Aspose.Words behavior, repetitive Microsoft Word 2016, use 'ContinuousSectionRestart.FROM_NEW_PAGE_ONLY'.
        # Page numbering restarts only if there is no other content before the section on the page where the section starts,
        # because of that the numbering will reset to 2 from the second page.
        doc.layout_options.continuous_section_page_numbering_restart = aw.layout.ContinuousSectionRestart.FROM_NEW_PAGE_ONLY
        doc.update_page_layout()

        doc.save(ARTIFACTS_DIR + "Layout.restart_page_numbering_in_continuous_section.pdf")
        #ExEnd
