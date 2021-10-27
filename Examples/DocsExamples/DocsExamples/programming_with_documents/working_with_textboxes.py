import unittest
import os
import sys

from docs_examples_base import DocsExamplesBase

import aspose.words as aw

class WorkingWithTextboxes(DocsExamplesBase):

    def test_create_a_link(self):

        #ExStart:CreateALink
        doc = aw.Document()

        shape1 = aw.drawing.Shape(doc, aw.drawing.ShapeType.TEXT_BOX)
        shape2 = aw.drawing.Shape(doc, aw.drawing.ShapeType.TEXT_BOX)

        text_box1 = shape1.text_box
        text_box2 = shape2.text_box

        if text_box1.is_valid_link_target(text_box2):
            text_box1.next = text_box2
        #ExEnd:CreateALink

    def test_check_sequence(self):

        #ExStart:CheckSequence
        doc = aw.Document()

        shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.TEXT_BOX)
        text_box = shape.text_box

        if text_box.next is not None and text_box.previous is None:
            print("The head of the sequence")

        if text_box.next is not None and text_box.previous is not None:
            print("The Middle of the sequence.")

        if text_box.next is None and text_box.previous is not None:
            print("The Tail of the sequence.")
        #ExEnd:CheckSequence

    def test_break_a_link(self):

        #ExStart:BreakALink
        doc = aw.Document()

        shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.TEXT_BOX)
        text_box = shape.text_box

        # Break a forward link.
        text_box.break_forward_link()

        # Break a forward link by setting a None.
        text_box.next = None

        # Break a link, which leads to this textbox.
        if text_box.previous is not None:
            text_box.previous.break_forward_link()
        #ExEnd:BreakALink
