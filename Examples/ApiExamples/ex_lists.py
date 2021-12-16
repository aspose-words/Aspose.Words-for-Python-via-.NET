# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import unittest
import io

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir, image_dir
from document_helper import DocumentHelper
from testutil import TestUtil

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir
IMAGE_DIR = image_dir

class ExLists(ApiExampleBase):

    def test_apply_default_bullets_and_numbers(self):

        #ExStart
        #ExFor:DocumentBuilder.list_format
        #ExFor:ListFormat.apply_number_default
        #ExFor:ListFormat.apply_bullet_default
        #ExFor:ListFormat.list_indent
        #ExFor:ListFormat.list_outdent
        #ExFor:ListFormat.remove_numbers
        #ExFor:ListFormat.list_level_number
        #ExSummary:Shows how to create bulleted and numbered lists.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Aspose.Words main advantages are:")

        # A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
        # We can create nested lists by increasing the indent level.
        # We can begin and end a list by using a document builder's "list_format" property.
        # Each paragraph that we add between a list's start and the end will become an item in the list.
        # Below are two types of lists that we can create with a document builder.
        # 1 -  A bulleted list:
        # This list will apply an indent and a bullet symbol ("•") before each paragraph.
        builder.list_format.apply_bullet_default()
        builder.writeln("Great performance")
        builder.writeln("High reliability")
        builder.writeln("Quality code and working")
        builder.writeln("Wide variety of features")
        builder.writeln("Easy to understand API")

        # End the bulleted list.
        builder.list_format.remove_numbers()

        builder.insert_break(aw.BreakType.PARAGRAPH_BREAK)
        builder.writeln("Aspose.Words allows:")

        # 2 -  A numbered list:
        # Numbered lists create a logical order for their paragraphs by numbering each item.
        builder.list_format.apply_number_default()

        # This paragraph is the first item. The first item of a numbered list will have a "1." as its list item symbol.
        builder.writeln("Opening documents from different formats:")

        self.assertEqual(0, builder.list_format.list_level_number)

        # Call the "list_indent" method to increase the current list level,
        # which will start a new self-contained list, with a deeper indent, at the current item of the first list level.
        builder.list_format.list_indent()

        self.assertEqual(1, builder.list_format.list_level_number)

        # These are the first three list items of the second list level, which will maintain a count
        # independent of the count of the first list level. According to the current list format,
        # they will have symbols of "a.", "b.", and "c.".
        builder.writeln("DOC")
        builder.writeln("PDF")
        builder.writeln("HTML")

        # Call the "list_outdent" method to return to the previous list level.
        builder.list_format.list_outdent()

        self.assertEqual(0, builder.list_format.list_level_number)

        # These two paragraphs will continue the count of the first list level.
        # These items will have symbols of "2.", and "3."
        builder.writeln("Processing documents")
        builder.writeln("Saving documents in different formats:")

        # If we increase the list level to a level that we have added items to previously,
        # the nested list will be separate from the previous, and its numbering will start from the beginning.
        # These list items will have symbols of "a.", "b.", "c.", "d.", and "e".
        builder.list_format.list_indent()
        builder.writeln("DOC")
        builder.writeln("PDF")
        builder.writeln("HTML")
        builder.writeln("MHTML")
        builder.writeln("Plain text")

        # Outdent the list level again.
        builder.list_format.list_outdent()
        builder.writeln("Doing many other things!")

        # End the numbered list.
        builder.list_format.remove_numbers()

        doc.save(ARTIFACTS_DIR + "Lists.apply_default_bullets_and_numbers.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Lists.apply_default_bullets_and_numbers.docx")

        TestUtil.verify_list_level(self, "\0.", 18.0, aw.NumberStyle.ARABIC, doc.lists[1].list_levels[0])
        TestUtil.verify_list_level(self, "\u0001.", 54.0, aw.NumberStyle.LOWERCASE_LETTER, doc.lists[1].list_levels[1])
        TestUtil.verify_list_level(self, "\uf0b7", 18.0, aw.NumberStyle.BULLET, doc.lists[0].list_levels[0])

    def test_specify_list_level(self):

        #ExStart
        #ExFor:ListCollection
        #ExFor:List
        #ExFor:ListFormat
        #ExFor:ListFormat.is_list_item
        #ExFor:ListFormat.list_level_number
        #ExFor:ListFormat.list
        #ExFor:ListTemplate
        #ExFor:DocumentBase.lists
        #ExFor:ListCollection.add(ListTemplate)
        #ExSummary:Shows how to work with list levels.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        self.assertFalse(builder.list_format.is_list_item)

        # A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
        # We can create nested lists by increasing the indent level.
        # We can begin and end a list by using a document builder's "list_format" property.
        # Each paragraph that we add between a list's start and the end will become an item in the list.
        # Below are two types of lists that we can create using a document builder.
        # 1 -  A numbered list:
        # Numbered lists create a logical order for their paragraphs by numbering each item.
        builder.list_format.list = doc.lists.add(aw.lists.ListTemplate.NUMBER_DEFAULT)

        self.assertTrue(builder.list_format.is_list_item)

        # By setting the "list_level_number" property, we can increase the list level
        # to begin a self-contained sub-list at the current list item.
        # The Microsoft Word list template called "NumberDefault" uses numbers to create list levels for the first list level.
        # Deeper list levels use letters and lowercase Roman numerals.
        for i in range(9):

            builder.list_format.list_level_number = i
            builder.writeln(f"Level {i}")

        # 2 -  A bulleted list:
        # This list will apply an indent and a bullet symbol ("•") before each paragraph.
        # Deeper levels of this list will use different symbols, such as "■" and "○".
        builder.list_format.list = doc.lists.add(aw.lists.ListTemplate.BULLET_DEFAULT)

        for i in range(9):
            builder.list_format.list_level_number = i
            builder.writeln(f"Level {i}")

        # We can disable list formatting to not format any subsequent paragraphs as lists by un-setting the "List" flag.
        builder.list_format.list = None

        self.assertFalse(builder.list_format.is_list_item)

        doc.save(ARTIFACTS_DIR + "Lists.specify_list_level.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Lists.specify_list_level.docx")

        TestUtil.verify_list_level(self, "\0.", 18.0, aw.NumberStyle.ARABIC, doc.lists[0].list_levels[0])

    def test_nested_lists(self):

        #ExStart
        #ExFor:ListFormat.list
        #ExFor:ParagraphFormat.clear_formatting
        #ExFor:ParagraphFormat.drop_cap_position
        #ExFor:ParagraphFormat.is_list_item
        #ExFor:Paragraph.is_list_item
        #ExSummary:Shows how to nest a list inside another list.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
        # We can create nested lists by increasing the indent level.
        # We can begin and end a list by using a document builder's "list_format" property.
        # Each paragraph that we add between a list's start and the end will become an item in the list.
        # Create an outline list for the headings.
        outline_list = doc.lists.add(aw.lists.ListTemplate.OUTLINE_NUMBERS)
        builder.list_format.list = outline_list
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1
        builder.writeln("This is my Chapter 1")

        # Create a numbered list.
        numbered_list = doc.lists.add(aw.lists.ListTemplate.NUMBER_DEFAULT)
        builder.list_format.list = numbered_list
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.NORMAL
        builder.writeln("Numbered list item 1.")

        # Every paragraph that comprises a list will have this flag.
        self.assertTrue(builder.current_paragraph.is_list_item)
        self.assertTrue(builder.paragraph_format.is_list_item)

        # Create a bulleted list.
        bulleted_list = doc.lists.add(aw.lists.ListTemplate.BULLET_DEFAULT)
        builder.list_format.list = bulleted_list
        builder.paragraph_format.left_indent = 72
        builder.writeln("Bulleted list item 1.")
        builder.writeln("Bulleted list item 2.")
        builder.paragraph_format.clear_formatting()

        # Revert to the numbered list.
        builder.list_format.list = numbered_list
        builder.writeln("Numbered list item 2.")
        builder.writeln("Numbered list item 3.")

        # Revert to the outline list.
        builder.list_format.list = outline_list
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1
        builder.writeln("This is my Chapter 2")

        builder.paragraph_format.clear_formatting()

        builder.document.save(ARTIFACTS_DIR + "Lists.nested_lists.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Lists.nested_lists.docx")

        TestUtil.verify_list_level(self, "\0)", 0.0, aw.NumberStyle.ARABIC, doc.lists[0].list_levels[0])
        TestUtil.verify_list_level(self, "\0.", 18.0, aw.NumberStyle.ARABIC, doc.lists[1].list_levels[0])
        TestUtil.verify_list_level(self, "\uf0b7", 18.0, aw.NumberStyle.BULLET, doc.lists[2].list_levels[0])

    def test_create_custom_list(self):

        #ExStart
        #ExFor:List
        #ExFor:List.list_levels
        #ExFor:ListFormat.list_level
        #ExFor:ListLevelCollection
        #ExFor:ListLevelCollection.__getitem__
        #ExFor:ListLevel
        #ExFor:ListLevel.alignment
        #ExFor:ListLevel.font
        #ExFor:ListLevel.number_style
        #ExFor:ListLevel.start_at
        #ExFor:ListLevel.trailing_character
        #ExFor:ListLevelAlignment
        #ExFor:NumberStyle
        #ExFor:ListTrailingCharacter
        #ExFor:ListLevel.number_format
        #ExFor:ListLevel.number_position
        #ExFor:ListLevel.text_position
        #ExFor:ListLevel.tab_position
        #ExSummary:Shows how to apply custom list formatting to paragraphs when using DocumentBuilder.
        doc = aw.Document()

        # A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
        # We can create nested lists by increasing the indent level.
        # We can begin and end a list by using a document builder's "list_format" property.
        # Each paragraph that we add between a list's start and the end will become an item in the list.
        # Create a list from a Microsoft Word template, and customize the first two of its list levels.
        list = doc.lists.add(aw.lists.ListTemplate.NUMBER_DEFAULT)

        list_level = list.list_levels[0]
        list_level.font.color = drawing.Color.red
        list_level.font.size = 24
        list_level.number_style = aw.NumberStyle.ORDINAL_TEXT
        list_level.start_at = 21
        list_level.number_format = "\x0000"

        list_level.number_position = -36
        list_level.text_position = 144
        list_level.tab_position = 144

        list_level = list.list_levels[1]
        list_level.alignment = aw.lists.ListLevelAlignment.RIGHT
        list_level.number_style = aw.NumberStyle.BULLET
        list_level.font.name = "Wingdings"
        list_level.font.color = drawing.Color.blue
        list_level.font.size = 24

        # This NumberFormat value will create star-shaped bullet list symbols.
        list_level.number_format = "\xf0af"
        list_level.trailing_character = aw.lists.ListTrailingCharacter.SPACE
        list_level.number_position = 144

        # Create paragraphs and apply both list levels of our custom list formatting to them.
        builder = aw.DocumentBuilder(doc)

        builder.list_format.list = list
        builder.writeln("The quick brown fox...")
        builder.writeln("The quick brown fox...")

        builder.list_format.list_indent()
        builder.writeln("jumped over the lazy dog.")
        builder.writeln("jumped over the lazy dog.")

        builder.list_format.list_outdent()
        builder.writeln("The quick brown fox...")

        builder.list_format.remove_numbers()

        builder.document.save(ARTIFACTS_DIR + "Lists.create_custom_list.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Lists.create_custom_list.docx")

        list_level = doc.lists[0].list_levels[0]

        TestUtil.verify_list_level(self, "\0", -36.0, aw.NumberStyle.ORDINAL_TEXT, list_level)
        self.assertEqual(drawing.Color.red.to_argb(), list_level.font.color.to_argb())
        self.assertEqual(24.0, list_level.font.size)
        self.assertEqual(21, list_level.start_at)

        list_level = doc.lists[0].list_levels[1]

        TestUtil.verify_list_level(self, "\xf0af", 144.0, aw.NumberStyle.BULLET, list_level)
        self.assertEqual(drawing.Color.blue.to_argb(), list_level.font.color.to_argb())
        self.assertEqual(24.0, list_level.font.size)
        self.assertEqual(1, list_level.start_at)
        self.assertEqual(aw.lists.ListTrailingCharacter.SPACE, list_level.trailing_character)

    def test_restart_numbering_using_list_copy(self):

        #ExStart
        #ExFor:List
        #ExFor:ListCollection
        #ExFor:ListCollection.add(ListTemplate)
        #ExFor:ListCollection.add_copy(List)
        #ExFor:ListLevel.start_at
        #ExFor:ListTemplate
        #ExSummary:Shows how to restart numbering in a list by copying a list.
        doc = aw.Document()

        # A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
        # We can create nested lists by increasing the indent level.
        # We can begin and end a list by using a document builder's "list_format" property.
        # Each paragraph that we add between a list's start and the end will become an item in the list.
        # Create a list from a Microsoft Word template, and customize its first list level.
        list1 = doc.lists.add(aw.lists.ListTemplate.NUMBER_ARABIC_PARENTHESIS)
        list1.list_levels[0].font.color = drawing.Color.red
        list1.list_levels[0].alignment = aw.lists.ListLevelAlignment.RIGHT

        # Apply our list to some paragraphs.
        builder = aw.DocumentBuilder(doc)

        builder.writeln("List 1 starts below:")
        builder.list_format.list = list1
        builder.writeln("Item 1")
        builder.writeln("Item 2")
        builder.list_format.remove_numbers()

        # We can add a copy of an existing list to the document's list collection
        # to create a similar list without making changes to the original.
        list2 = doc.lists.add_copy(list1)
        list2.list_levels[0].font.color = drawing.Color.blue
        list2.list_levels[0].start_at = 10

        # Apply the second list to new paragraphs.
        builder.writeln("List 2 starts below:")
        builder.list_format.list = list2
        builder.writeln("Item 1")
        builder.writeln("Item 2")
        builder.list_format.remove_numbers()

        doc.save(ARTIFACTS_DIR + "Lists.restart_numbering_using_list_copy.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Lists.restart_numbering_using_list_copy.docx")

        list1 = doc.lists[0]
        TestUtil.verify_list_level(self, "\0)", 18.0, aw.NumberStyle.ARABIC, list1.list_levels[0])
        self.assertEqual(drawing.Color.red.to_argb(), list1.list_levels[0].font.color.to_argb())
        self.assertEqual(10.0, list1.list_levels[0].font.size)
        self.assertEqual(1, list1.list_levels[0].start_at)

        list2 = doc.lists[1]
        TestUtil.verify_list_level(self, "\0)", 18.0, aw.NumberStyle.ARABIC, list2.list_levels[0])
        self.assertEqual(drawing.Color.blue.to_argb(), list2.list_levels[0].font.color.to_argb())
        self.assertEqual(10.0, list2.list_levels[0].font.size)
        self.assertEqual(10, list2.list_levels[0].start_at)

    def test_create_and_use_list_style(self):

        #ExStart
        #ExFor:StyleCollection.add(StyleType,str)
        #ExFor:Style.list
        #ExFor:StyleType
        #ExFor:List.is_list_style_definition
        #ExFor:List.is_list_style_reference
        #ExFor:List.is_multi_level
        #ExFor:List.style
        #ExFor:ListLevelCollection
        #ExFor:ListLevelCollection.count
        #ExFor:ListLevelCollection.__getitem__
        #ExFor:ListCollection.add(Style)
        #ExSummary:Shows how to create a list style and use it in a document.
        doc = aw.Document()

        # A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
        # We can create nested lists by increasing the indent level.
        # We can begin and end a list by using a document builder's "list_format" property.
        # Each paragraph that we add between a list's start and the end will become an item in the list.
        # We can contain an entire List object within a style.
        list_style = doc.styles.add(aw.StyleType.LIST, "MyListStyle")

        list1 = list_style.list

        self.assertTrue(list1.is_list_style_definition)
        self.assertFalse(list1.is_list_style_reference)
        self.assertTrue(list1.is_multi_level)
        self.assertEqual(list_style, list1.style)

        # Change the appearance of all list levels in our list.
        for level in list1.list_levels:
            level.font.name = "Verdana"
            level.font.color = drawing.Color.blue
            level.font.bold = True

        builder = aw.DocumentBuilder(doc)

        builder.writeln("Using list style first time:")

        # Create another list from a list within a style.
        list2 = doc.lists.add(list_style)

        self.assertFalse(list2.is_list_style_definition)
        self.assertTrue(list2.is_list_style_reference)
        self.assertEqual(list_style, list2.style)

        # Add some list items that our list will format.
        builder.list_format.list = list2
        builder.writeln("Item 1")
        builder.writeln("Item 2")
        builder.list_format.remove_numbers()

        builder.writeln("Using list style second time:")

        # Create and apply another list based on the list style.
        list3 = doc.lists.add(list_style)
        builder.list_format.list = list3
        builder.writeln("Item 1")
        builder.writeln("Item 2")
        builder.list_format.remove_numbers()

        builder.document.save(ARTIFACTS_DIR + "Lists.create_and_use_list_style.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Lists.create_and_use_list_style.docx")

        list1 = doc.lists[0]

        TestUtil.verify_list_level(self, "\0.", 18.0, aw.NumberStyle.ARABIC, list1.list_levels[0])
        self.assertTrue(list1.is_list_style_definition)
        self.assertFalse(list1.is_list_style_reference)
        self.assertTrue(list1.is_multi_level)
        self.assertEqual(drawing.Color.blue.to_argb(), list1.list_levels[0].font.color.to_argb())
        self.assertEqual("Verdana", list1.list_levels[0].font.name)
        self.assertTrue(list1.list_levels[0].font.bold)

        list2 = doc.lists[1]

        TestUtil.verify_list_level(self, "\0.", 18.0, aw.NumberStyle.ARABIC, list2.list_levels[0])
        self.assertFalse(list2.is_list_style_definition)
        self.assertTrue(list2.is_list_style_reference)
        self.assertTrue(list2.is_multi_level)

        list3 = doc.lists[2]

        TestUtil.verify_list_level(self, "\0.", 18.0, aw.NumberStyle.ARABIC, list3.list_levels[0])
        self.assertFalse(list3.is_list_style_definition)
        self.assertTrue(list3.is_list_style_reference)
        self.assertTrue(list3.is_multi_level)

    def test_detect_bulleted_paragraphs(self):

        #ExStart
        #ExFor:Paragraph.list_format
        #ExFor:ListFormat.is_list_item
        #ExFor:CompositeNode.get_text
        #ExFor:List.list_id
        #ExSummary:Shows how to output all paragraphs in a document that are list items.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.list_format.apply_number_default()
        builder.writeln("Numbered list item 1")
        builder.writeln("Numbered list item 2")
        builder.writeln("Numbered list item 3")
        builder.list_format.remove_numbers()

        builder.list_format.apply_bullet_default()
        builder.writeln("Bulleted list item 1")
        builder.writeln("Bulleted list item 2")
        builder.writeln("Bulleted list item 3")
        builder.list_format.remove_numbers()

        paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)

        for para in paras:
            para = para.as_paragraph()
            if para.list_format.is_list_item:
                print(f"This paragraph belongs to list ID# {para.list_format.list.list_id}, number style \"{para.list_format.list_level.number_style}\"")
                print(f"\t\"{para.get_text().strip()}\"")

        #ExEnd

        doc = DocumentHelper.save_open(doc)
        paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)

        self.assertEqual(6, len([p for p in paras if p.as_paragraph().list_format.is_list_item]))

    def test_remove_bullets_from_paragraphs(self):

        #ExStart
        #ExFor:ListFormat.remove_numbers
        #ExSummary:Shows how to remove list formatting from all paragraphs in the main text of a section.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.list_format.apply_number_default()
        builder.writeln("Numbered list item 1")
        builder.writeln("Numbered list item 2")
        builder.writeln("Numbered list item 3")
        builder.list_format.remove_numbers()

        paras = [node.as_paragraph() for node in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)]

        self.assertEqual(3, len([p for p in paras if p.list_format.is_list_item]))

        for paragraph in paras:
            paragraph.list_format.remove_numbers()

        self.assertEqual(0, len([p for p in paras if p.list_format.is_list_item]))
        #ExEnd

    def test_apply_existing_list_to_paragraphs(self):

        #ExStart
        #ExFor:ListCollection.__getitem__(int)
        #ExSummary:Shows how to apply list formatting of an existing list to a collection of paragraphs.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Paragraph 1")
        builder.writeln("Paragraph 2")
        builder.write("Paragraph 3")

        paras = [node.as_paragraph() for node in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)]

        self.assertEqual(0, len([p for p in paras if p.list_format.is_list_item]))

        doc.lists.add(aw.lists.ListTemplate.NUMBER_DEFAULT)
        list = doc.lists[0]

        for paragraph in paras:
            paragraph.list_format.list = list
            paragraph.list_format.list_level_number = 2

        self.assertEqual(3, len([p for p in paras if p.list_format.is_list_item]))
        #ExEnd

        doc = DocumentHelper.save_open(doc)
        paras = [node.as_paragraph() for node in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)]

        self.assertEqual(3, len([p for p in paras if p.list_format.is_list_item]))
        self.assertEqual(3, len([p for p in paras if p.list_format.list_level_number == 2]))

    def test_apply_new_list_to_paragraphs(self):

        #ExStart
        #ExFor:ListCollection.add(ListTemplate)
        #ExSummary:Shows how to create a list by applying a new list format to a collection of paragraphs.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Paragraph 1")
        builder.writeln("Paragraph 2")
        builder.write("Paragraph 3")

        paras = [node.as_paragraph() for node in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)]

        self.assertEqual(0, len([p for p in paras if p.list_format.is_list_item]))

        list = doc.lists.add(aw.lists.ListTemplate.NUMBER_UPPERCASE_LETTER_DOT)

        for paragraph in paras:
            paragraph.list_format.list = list
            paragraph.list_format.list_level_number = 1

        self.assertEqual(3, len([p for p in paras if p.list_format.is_list_item]))
        #ExEnd

        doc = DocumentHelper.save_open(doc)
        paras = [node.as_paragraph() for node in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)]

        self.assertEqual(3, len([p for p in paras if p.list_format.is_list_item]))
        self.assertEqual(3, len([p for p in paras if p.list_format.list_level_number == 1]))

    #ExStart
    #ExFor:ListTemplate
    #ExSummary:Shows how to create a document that contains all outline headings list templates.
    def test_outline_heading_templates(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        list = doc.lists.add(aw.lists.ListTemplate.OUTLINE_HEADINGS_ARTICLE_SECTION)
        ExLists.add_outline_heading_paragraphs(builder, list, "Aspose.Words Outline - \"Article Section\"")

        list = doc.lists.add(aw.lists.ListTemplate.OUTLINE_HEADINGS_LEGAL)
        ExLists.add_outline_heading_paragraphs(builder, list, "Aspose.Words Outline - \"Legal\"")

        builder.insert_break(aw.BreakType.PAGE_BREAK)

        list = doc.lists.add(aw.lists.ListTemplate.OUTLINE_HEADINGS_NUMBERS)
        ExLists.add_outline_heading_paragraphs(builder, list, "Aspose.Words Outline - \"Numbers\"")

        list = doc.lists.add(aw.lists.ListTemplate.OUTLINE_HEADINGS_CHAPTER)
        ExLists.add_outline_heading_paragraphs(builder, list, "Aspose.Words Outline - \"Chapters\"")

        doc.save(ARTIFACTS_DIR + "Lists.outline_heading_templates.docx")
        self._test_outline_heading_templates(aw.Document(ARTIFACTS_DIR + "Lists.outline_heading_templates.docx")); #ExSkip

    @staticmethod
    def add_outline_heading_paragraphs(builder: aw.DocumentBuilder, list: aw.lists.List, title: str):

        builder.paragraph_format.clear_formatting()
        builder.writeln(title)

        for i in range(9):
            builder.list_format.list = list
            builder.list_format.list_level_number = i

            style_name = f"Heading {i + 1}"
            builder.paragraph_format.style_name = style_name
            builder.writeln(style_name)

        builder.list_format.remove_numbers()

    #ExEnd

    def _test_outline_heading_templates(self, doc: aw.Document):

        list = doc.lists[0] # Article section list template.

        TestUtil.verify_list_level(self, "Article \0.", 0.0, aw.NumberStyle.UPPERCASE_ROMAN, list.list_levels[0])
        TestUtil.verify_list_level(self, "Section \0.\u0001", 0.0, aw.NumberStyle.LEADING_ZERO, list.list_levels[1])
        TestUtil.verify_list_level(self, "(\u0002)", 14.4, aw.NumberStyle.LOWERCASE_LETTER, list.list_levels[2])
        TestUtil.verify_list_level(self, "(\u0003)", 36.0, aw.NumberStyle.LOWERCASE_ROMAN, list.list_levels[3])
        TestUtil.verify_list_level(self, "\u0004)", 28.8, aw.NumberStyle.ARABIC, list.list_levels[4])
        TestUtil.verify_list_level(self, "\u0005)", 36.0, aw.NumberStyle.LOWERCASE_LETTER, list.list_levels[5])
        TestUtil.verify_list_level(self, "\u0006)", 50.4, aw.NumberStyle.LOWERCASE_ROMAN, list.list_levels[6])
        TestUtil.verify_list_level(self, "\a.", 50.4, aw.NumberStyle.LOWERCASE_LETTER, list.list_levels[7])
        TestUtil.verify_list_level(self, "\b.", 72.0, aw.NumberStyle.LOWERCASE_ROMAN, list.list_levels[8])

        list = doc.lists[1] # Legal list template.

        TestUtil.verify_list_level(self, "\0", 0.0, aw.NumberStyle.ARABIC, list.list_levels[0])
        TestUtil.verify_list_level(self, "\0.\u0001", 0.0, aw.NumberStyle.ARABIC, list.list_levels[1])
        TestUtil.verify_list_level(self, "\0.\u0001.\u0002", 0.0, aw.NumberStyle.ARABIC, list.list_levels[2])
        TestUtil.verify_list_level(self, "\0.\u0001.\u0002.\u0003", 0.0, aw.NumberStyle.ARABIC, list.list_levels[3])
        TestUtil.verify_list_level(self, "\0.\u0001.\u0002.\u0003.\u0004", 0.0, aw.NumberStyle.ARABIC, list.list_levels[4])
        TestUtil.verify_list_level(self, "\0.\u0001.\u0002.\u0003.\u0004.\u0005", 0.0, aw.NumberStyle.ARABIC, list.list_levels[5])
        TestUtil.verify_list_level(self, "\0.\u0001.\u0002.\u0003.\u0004.\u0005.\u0006", 0.0, aw.NumberStyle.ARABIC, list.list_levels[6])
        TestUtil.verify_list_level(self, "\0.\u0001.\u0002.\u0003.\u0004.\u0005.\u0006.\a", 0.0, aw.NumberStyle.ARABIC, list.list_levels[7])
        TestUtil.verify_list_level(self, "\0.\u0001.\u0002.\u0003.\u0004.\u0005.\u0006.\a.\b", 0.0, aw.NumberStyle.ARABIC, list.list_levels[8])

        list = doc.lists[2] # Numbered list template.

        TestUtil.verify_list_level(self, "\0.", 0.0, aw.NumberStyle.UPPERCASE_ROMAN, list.list_levels[0])
        TestUtil.verify_list_level(self, "\u0001.", 36.0, aw.NumberStyle.UPPERCASE_LETTER, list.list_levels[1])
        TestUtil.verify_list_level(self, "\u0002.", 72.0, aw.NumberStyle.ARABIC, list.list_levels[2])
        TestUtil.verify_list_level(self, "\u0003)", 108.0, aw.NumberStyle.LOWERCASE_LETTER, list.list_levels[3])
        TestUtil.verify_list_level(self, "(\u0004)", 144.0, aw.NumberStyle.ARABIC, list.list_levels[4])
        TestUtil.verify_list_level(self, "(\u0005)", 180.0, aw.NumberStyle.LOWERCASE_LETTER, list.list_levels[5])
        TestUtil.verify_list_level(self, "(\u0006)", 216.0, aw.NumberStyle.LOWERCASE_ROMAN, list.list_levels[6])
        TestUtil.verify_list_level(self, "(\a)", 252.0, aw.NumberStyle.LOWERCASE_LETTER, list.list_levels[7])
        TestUtil.verify_list_level(self, "(\b)", 288.0, aw.NumberStyle.LOWERCASE_ROMAN, list.list_levels[8])

        list = doc.lists[3] # Chapter list template.

        TestUtil.verify_list_level(self, "Chapter \0", 0.0, aw.NumberStyle.ARABIC, list.list_levels[0])
        TestUtil.verify_list_level(self, "", 0.0, aw.NumberStyle.NONE, list.list_levels[1])
        TestUtil.verify_list_level(self, "", 0.0, aw.NumberStyle.NONE, list.list_levels[2])
        TestUtil.verify_list_level(self, "", 0.0, aw.NumberStyle.NONE, list.list_levels[3])
        TestUtil.verify_list_level(self, "", 0.0, aw.NumberStyle.NONE, list.list_levels[4])
        TestUtil.verify_list_level(self, "", 0.0, aw.NumberStyle.NONE, list.list_levels[5])
        TestUtil.verify_list_level(self, "", 0.0, aw.NumberStyle.NONE, list.list_levels[6])
        TestUtil.verify_list_level(self, "", 0.0, aw.NumberStyle.NONE, list.list_levels[7])
        TestUtil.verify_list_level(self, "", 0.0, aw.NumberStyle.NONE, list.list_levels[8])

    #ExStart
    #ExFor:ListCollection
    #ExFor:ListCollection.add_copy(List)
    #ExFor:ListCollection.__iter__
    #ExSummary:Shows how to create a document with a sample of all the lists from another document.
    def test_print_out_all_lists(self):

        src_doc = aw.Document(MY_DIR + "Rendering.docx")

        dst_doc = aw.Document()
        builder = aw.DocumentBuilder(dst_doc)

        for src_list in src_doc.lists:
            dst_list = dst_doc.lists.add_copy(src_list)
            ExLists.add_list_sample(builder, dst_list)

        dst_doc.save(ARTIFACTS_DIR + "Lists.print_out_all_lists.docx")
        self._test_print_out_all_ists(src_doc, aw.Document(ARTIFACTS_DIR + "Lists.print_out_all_lists.docx")) #ExSkip

    @staticmethod
    def add_list_sample(builder: aw.DocumentBuilder, list: aw.lists.List):

        builder.writeln(f"Sample formatting of list with list_id: {list.list_id}")
        builder.list_format.list = list
        for i in range(list.list_levels.count):
            builder.list_format.list_level_number = i
            builder.writeln(f"Level {i}")

        builder.list_format.remove_numbers()
        builder.writeln()

    #ExEnd

    def _test_print_out_all_ists(self, list_source_doc: aw.Document, out_doc: aw.Document):

        for list in out_doc.lists:
            for i in range(list.list_levels.count):
                for l in list_source_doc.lists:
                    if l.list_id == list.list_id:
                        expected_list_level = l.list_levels[i]
                        self.assertEqual(expected_list_level.number_format, list.list_levels[i].number_format)
                        self.assertEqual(expected_list_level.number_position, list.list_levels[i].number_position)
                        self.assertEqual(expected_list_level.number_style, list.list_levels[i].number_style)
                        break

    def test_list_document(self):

        #ExStart
        #ExFor:ListCollection.document
        #ExFor:ListCollection.count
        #ExFor:ListCollection.__getitem__(int)
        #ExFor:ListCollection.get_list_by_list_id
        #ExFor:List.document
        #ExFor:List.list_id
        #ExSummary:Shows how to verify owner document properties of lists.
        doc = aw.Document()

        lists = doc.lists

        self.assertEqual(doc, lists.document)

        list = lists.add(aw.lists.ListTemplate.BULLET_DEFAULT)

        self.assertEqual(doc, list.document)

        print("Current list count:", lists.count)
        print("Is the first document list:", lists[0] is list)
        print("List id:", list.list_id)
        print("List is the same by list_id:", lists.get_list_by_list_id(1) is list)
        #ExEnd

        doc = DocumentHelper.save_open(doc)
        lists = doc.lists

        self.assertEqual(doc, lists.document)
        self.assertEqual(1, lists.count)
        self.assertEqual(1, lists[0].list_id)
        self.assertEqual(lists[0], lists.get_list_by_list_id(1))

    def test_create_list_restart_after_higher(self):

        #ExStart
        #ExFor:ListLevel.number_style
        #ExFor:ListLevel.number_format
        #ExFor:ListLevel.is_legal
        #ExFor:ListLevel.restart_after_level
        #ExFor:ListLevel.linked_style
        #ExFor:ListLevelCollection.__iter__
        #ExSummary:Shows advances ways of customizing list labels.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # A list allows us to organize and decorate sets of paragraphs with prefix symbols and indents.
        # We can create nested lists by increasing the indent level.
        # We can begin and end a list by using a document builder's "list_format" property.
        # Each paragraph that we add between a list's start and the end will become an item in the list.
        list = doc.lists.add(aw.lists.ListTemplate.NUMBER_DEFAULT)

        # Level 1 labels will be formatted according to the "Heading 1" paragraph style and will have a prefix.
        # These will look like "Appendix A", "Appendix B"...
        list.list_levels[0].number_format = "Appendix \x0000"
        list.list_levels[0].number_style = aw.NumberStyle.UPPERCASE_LETTER
        list.list_levels[0].linked_style = doc.styles.get_by_name("Heading 1")

        # Level 2 labels will display the current numbers of the first and the second list levels and have leading zeroes.
        # If the first list level is at 1, then the list labels from these will look like "Section (1.01)", "Section (1.02)"...
        list.list_levels[1].number_format = "Section (\x0000.\x0001)"
        list.list_levels[1].number_style = aw.NumberStyle.LEADING_ZERO

        # Note that the higher-level uses UppercaseLetter numbering.
        # We can set the "is_legal" property to use Arabic numbers for the higher list levels.
        list.list_levels[1].is_legal = True
        list.list_levels[1].restart_after_level = 0

        # Level 3 labels will be upper case Roman numerals with a prefix and a suffix and will restart at each List level 1 item.
        # These list labels will look like "-I-", "-II-"...
        list.list_levels[2].number_format = "-\x0002-"
        list.list_levels[2].number_style = aw.NumberStyle.UPPERCASE_ROMAN
        list.list_levels[2].restart_after_level = 1

        # Make labels of all list levels bold.
        for level in list.list_levels:
            level.font.bold = True

        # Apply list formatting to the current paragraph.
        builder.list_format.list = list

        # Create list items that will display all three of our list levels.
        for n in range(2):
            for i in range(3):
                builder.list_format.list_level_number = i
                builder.writeln(f"Level {i}")

        builder.list_format.remove_numbers()

        doc.save(ARTIFACTS_DIR + "Lists.create_list_restart_after_higher.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Lists.create_list_restart_after_higher.docx")

        list_level = doc.lists[0].list_levels[0]

        TestUtil.verify_list_level(self, "Appendix \0", 18.0, aw.NumberStyle.UPPERCASE_LETTER, list_level)
        self.assertFalse(list_level.is_legal)
        self.assertEqual(-1, list_level.restart_after_level)
        self.assertEqual("Heading 1", list_level.linked_style.name)

        list_level = doc.lists[0].list_levels[1]

        TestUtil.verify_list_level(self, "Section (\0.\u0001)", 54.0, aw.NumberStyle.LEADING_ZERO, list_level)
        self.assertTrue(list_level.is_legal)
        self.assertEqual(0, list_level.restart_after_level)
        self.assertIsNone(list_level.linked_style)

    def test_get_list_labels(self):

        #ExStart
        #ExFor:Document.update_list_labels()
        #ExFor:Node.to_string(SaveFormat)
        #ExFor:ListLabel
        #ExFor:Paragraph.list_label
        #ExFor:ListLabel.label_value
        #ExFor:ListLabel.label_string
        #ExSummary:Shows how to extract the list labels of all paragraphs that are list items.
        doc = aw.Document(MY_DIR + "Rendering.docx")
        doc.update_list_labels()

        paras = [node.as_paragraph() for node in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)]

        # Find if we have the paragraph list. In our document, our list uses plain Arabic numbers,
        # which start at three and ends at six.
        for paragraph in paras:
            if paragraph.list_format.is_list_item:
                print(f"List item paragraph #{paras.index(paragraph)}")

                # This is the text we get when getting when we output this node to text format.
                # This text output will omit list labels. Strip any paragraph formatting characters.
                paragraph_text = paragraph.to_string(aw.SaveFormat.TEXT).strip()
                print(f"\tExported Text: {paragraph_text}")

                label = paragraph.list_label

                # This gets the position of the paragraph in the current level of the list. If we have a list with multiple levels,
                # this will tell us what position it is on that level.
                print(f"\tNumerical Id: {label.label_value}")

                # Combine them together to include the list label with the text in the output.
                print(f"\tList label combined with text: {label.label_string} {paragraph_text}")

        #ExEnd

        self.assertEqual(10, len([p for p in paras if p.list_format.is_list_item]))

    def test_create_picture_bullet(self):

        #ExStart
        #ExFor:ListLevel.create_picture_bullet
        #ExFor:ListLevel.delete_picture_bullet
        #ExSummary:Shows how to set a custom image icon for list item labels.
        doc = aw.Document()

        list = doc.lists.add(aw.lists.ListTemplate.BULLET_CIRCLE)

        # Create a picture bullet for the current list level, and set an image from a local file system
        # as the icon that the bullets for this list level will display.
        list.list_levels[0].create_picture_bullet()
        list.list_levels[0].image_data.set_image(IMAGE_DIR + "Logo icon.ico")

        self.assertTrue(list.list_levels[0].image_data.has_image)

        builder = aw.DocumentBuilder(doc)

        builder.list_format.list = list
        builder.writeln("Hello world!")
        builder.write("Hello again!")

        doc.save(ARTIFACTS_DIR + "Lists.create_picture_bullet.docx")

        list.list_levels[0].delete_picture_bullet()

        self.assertIsNone(list.list_levels[0].image_data)
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Lists.create_picture_bullet.docx")

        self.assertTrue(doc.lists[0].list_levels[0].image_data.has_image)

    def test_custom_number_style_format(self):

        #ExStart
        #ExFor:ListLevel.custom_number_style_format
        #ExFor:ListLevel.get_effective_value(int,NumberStyle,str)
        #ExSummary:Shows how to get the format for a list with the custom number style.
        doc = aw.Document(MY_DIR + "List with leading zero.docx")

        list_level = doc.first_section.body.paragraphs[0].list_format.list_level

        custom_number_style_format = ""

        if list_level.number_style == aw.NumberStyle.CUSTOM:
            custom_number_style_format = list_level.custom_number_style_format

        self.assertEqual("001, 002, 003, ...", custom_number_style_format)

        # We can get value for the specified index of the list item.
        self.assertEqual("iv", list_level.get_effective_value(4, aw.NumberStyle.LOWERCASE_ROMAN, None))
        self.assertEqual("005", list_level.get_effective_value(5, aw.NumberStyle.CUSTOM, custom_number_style_format))
        #ExEnd

        with self.assertRaises(Exception):
            aw.lists.ListLevel.get_effective_value(5, aw.NumberStyle.LOWERCASE_ROMAN, custom_number_style_format)
        with self.assertRaises(Exception):
            aw.lists.ListLevel.get_effective_value(5, aw.NumberStyle.CUSTOM, None)
        with self.assertRaises(Exception):
            aw.lists.ListLevel.get_effective_value(5, aw.NumberStyle.CUSTOM, "....")
