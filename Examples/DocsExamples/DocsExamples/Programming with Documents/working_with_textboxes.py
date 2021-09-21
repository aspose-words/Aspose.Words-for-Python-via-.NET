import unittest
import os
import sys

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw

class WorkingWithTextboxes(docs_base.DocsExamplesBase):

    def test_create_a_link(self) :
        
        #ExStart:CreateALink
        doc = aw.Document()

        shape1 = aw.drawing.Shape(doc, aw.drawing.ShapeType.TEXT_BOX)
        shape2 = aw.drawing.Shape(doc, aw.drawing.ShapeType.TEXT_BOX)

        textBox1 = shape1.text_box
        textBox2 = shape2.text_box

        if textBox1.is_valid_link_target(textBox2) :
            textBox1.next = textBox2
        #ExEnd:CreateALink
        

    def test_check_sequence(self) :
        
        #ExStart:CheckSequence
        doc = aw.Document()

        shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.TEXT_BOX)
        textBox = shape.text_box

        if (textBox.next != None and textBox.previous == None) :
            print("The head of the sequence")
            

        if (textBox.next != None and textBox.previous != None) :
            print("The Middle of the sequence.")
            

        if (textBox.next == None and textBox.previous != None) :
            print("The Tail of the sequence.")
            
        #ExEnd:CheckSequence
        

    def test_break_a_link(self) :
        
        #ExStart:BreakALink
        doc = aw.Document()

        shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.TEXT_BOX)
        textBox = shape.text_box

        # Break a forward link.
        textBox.break_forward_link()

        # Break a forward link by setting a None.
        textBox.next = None

        # Break a link, which leads to this textbox.
        if textBox.previous != None :
            textBox.previous.break_forward_link()
        #ExEnd:BreakALink
        
    

if __name__ == '__main__':
    unittest.main()