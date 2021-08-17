import unittest

import api_example_base as aeb
from document_helper import DocumentHelper

import aspose.words as aw

class ExControlChar(aeb.ApiExampleBase):
    
    def test_carriage_return(self) :
        
        #ExStart
        #ExFor:ControlChar
        #ExFor:ControlChar.CR
        #ExFor:Node.get_text
        #ExSummary:Shows how to use control characters.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert paragraphs with text with DocumentBuilder.
        builder.writeln("Hello world!")
        builder.writeln("Hello again!")

        # Converting the document to text form reveals that control characters
        # represent some of the document's structural elements, such as page breaks.
        self.assertEqual(f"Hello world!{aw.ControlChar.CR}" +
                        f"Hello again!{aw.ControlChar.CR}" +
                        aw.ControlChar.PAGE_BREAK, doc.get_text())

        # When converting a document to string form,
        # we can omit some of the control characters with the Trim method.
        self.assertEqual(f"Hello world!{aw.ControlChar.CR}" +
                        "Hello again!", doc.get_text().strip())
        #ExEnd
        

    def test_insert_control_chars(self) :
        
        #ExStart
        #ExFor:.ControlChar.CELL
        #ExFor:.ControlChar.COLUMN_BREAK
        #ExFor:ControlChar.CR_LF
        #ExFor:ControlChar.LF
        #ExFor:ControlChar.LINE_BREAK
        #ExFor:ControlChar.LINE_FEED
        #ExFor:ControlChar.NON_BREAKING_SPACE
        #ExFor:.ControlChar.PAGE_BREAK
        #ExFor:ControlChar.PARAGRAPH_BREAK
        #ExFor:ControlChar.SECTION_BREAK
        #ExFor:ControlChar.cell_char
        #ExFor:ControlChar.column_break_char
        #ExFor:ControlChar.default_text_input_char
        #ExFor:ControlChar.field_end_char
        #ExFor:ControlChar.field_start_char
        #ExFor:ControlChar.field_separator_char
        #ExFor:ControlChar.line_break_char
        #ExFor:ControlChar.line_feed_char
        #ExFor:ControlChar.non_breaking_hyphen_char
        #ExFor:ControlChar.non_breaking_space_char
        #ExFor:ControlChar.optional_hyphen_char
        #ExFor:.ControlChar.PAGE_BREAK_char
        #ExFor:ControlChar.paragraph_break_char
        #ExFor:ControlChar.section_break_char
        #ExFor:ControlChar.space_char
        #ExSummary:Shows how to add various control characters to a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Add a regular space.
#        builder.write("Before space." + aw.ControlChar.SPACE_CHAR + "After space.")

        # Add an NBSP, which is a non-breaking space.
        # Unlike the regular space, this space cannot have an automatic line break at its position.
        builder.write("Before space." + aw.ControlChar.NON_BREAKING_SPACE + "After space.")

        # Add a tab character.
        builder.write("Before tab." + aw.ControlChar.TAB + "After tab.")

        # Add a line break.
        builder.write("Before line break." + aw.ControlChar.LINE_BREAK + "After line break.")

        # Add a new line and starts a new paragraph.
        self.assertEqual(1, doc.first_section.body.get_child_nodes(aw.NodeType.PARAGRAPH, True).count)
        builder.write("Before line feed." + aw.ControlChar.LINE_FEED + "After line feed.")
        self.assertEqual(2, doc.first_section.body.get_child_nodes(aw.NodeType.PARAGRAPH, True).count)

        # The line feed character has two versions.
        self.assertEqual(aw.ControlChar.LINE_FEED, aw.ControlChar.LF)

        # Carriage returns and line feeds can be represented together by one character.
        self.assertEqual(aw.ControlChar.CR_LF, aw.ControlChar.CR + aw.ControlChar.LF)

        # Add a paragraph break, which will start a new paragraph.
        builder.write("Before paragraph break." + aw.ControlChar.PARAGRAPH_BREAK + "After paragraph break.")
        self.assertEqual(3, doc.first_section.body.get_child_nodes(aw.NodeType.PARAGRAPH, True).count)

        # Add a section break. This does not make a new section or paragraph.
        self.assertEqual(1, doc.sections.count)
        builder.write("Before section break." + aw.ControlChar.SECTION_BREAK + "After section break.")
        self.assertEqual(1, doc.sections.count)

        # Add a page break.
        builder.write("Before page break." + aw.ControlChar.PAGE_BREAK + "After page break.")

        # A page break is the same value as a section break.
        self.assertEqual(aw.ControlChar.PAGE_BREAK, aw.ControlChar.SECTION_BREAK)

        # Insert a new section, and then set its column count to two.
        doc.append_child(aw.Section(doc))
        builder.move_to_section(1)
        builder.current_section.page_setup.text_columns.set_count(2)

        # We can use a control character to mark the point where text moves to the next column.
        builder.write("Text at end of column 1." + aw.ControlChar.COLUMN_BREAK + "Text at beginning of column 2.")

        doc.save(aeb.artifacts_dir + "ControlChar.insert_control_chars.docx")

        # There are char and string counterparts for most characters.
#        self.assertEqual(Convert.to_char(.ControlChar.CELL), ControlChar.cell_char)
#        self.assertEqual(Convert.to_char(ControlChar.NON_BREAKING_SPACE), ControlChar.non_breaking_space_char)
#        self.assertEqual(Convert.to_char(ControlChar.tab), ControlChar.tab_char)
#        self.assertEqual(Convert.to_char(ControlChar.LINE_BREAK), ControlChar.line_break_char)
#        self.assertEqual(Convert.to_char(ControlChar.LINE_FEED), ControlChar.line_feed_char)
#        self.assertEqual(Convert.to_char(ControlChar.PARAGRAPH_BREAK), ControlChar.paragraph_break_char)
#        self.assertEqual(Convert.to_char(ControlChar.SECTION_BREAK), ControlChar.section_break_char)
#        self.assertEqual(Convert.to_char(.ControlChar.PAGE_BREAK), ControlChar.section_break_char)
#        self.assertEqual(Convert.to_char(.ControlChar.COLUMN_BREAK), ControlChar.column_break_char)
        #ExEnd
        
    
if __name__ == '__main__':
    unittest.main()    
