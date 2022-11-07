# Copyright (c) 2001-2022 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import aspose.words as aw

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR
from document_helper import DocumentHelper

class ExParagraphFormat(ApiExampleBase):

    def test_asian_typography_properties(self):

        #ExStart
        #ExFor:ParagraphFormat.far_east_line_break_control
        #ExFor:ParagraphFormat.word_wrap
        #ExFor:ParagraphFormat.hanging_punctuation
        #ExSummary:Shows how to set special properties for Asian typography.
        doc = aw.Document(MY_DIR + "Document.docx")

        format = doc.first_section.body.first_paragraph.paragraph_format
        format.far_east_line_break_control = True
        format.word_wrap = False
        format.hanging_punctuation = True

        doc.save(ARTIFACTS_DIR + "ParagraphFormat.asian_typography_properties.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "ParagraphFormat.asian_typography_properties.docx")
        format = doc.first_section.body.first_paragraph.paragraph_format

        self.assertTrue(format.far_east_line_break_control)
        self.assertFalse(format.word_wrap)
        self.assertTrue(format.hanging_punctuation)

    def test_drop_cap(self):

        for drop_cap_position in (aw.DropCapPosition.MARGIN,
                                  aw.DropCapPosition.NORMAL,
                                  aw.DropCapPosition.NONE):
            with self.subTest():
                #ExStart
                #ExFor:DropCapPosition
                #ExSummary:Shows how to create a drop cap.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Insert one paragraph with a large letter that the text in the second and third paragraphs begins with.
                builder.font.size = 54
                builder.writeln("L")

                builder.font.size = 18
                builder.writeln("orem ipsum dolor sit amet, consectetur adipiscing elit, " +
                                "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ")
                builder.writeln("Ut enim ad minim veniam, quis nostrud exercitation " +
                                "ullamco laboris nisi ut aliquip ex ea commodo consequat.")

                # Currently, the second and third paragraphs will appear underneath the first.
                # We can convert the first paragraph as a drop cap for the other paragraphs via its "ParagraphFormat" object.
                # Set the "drop_cap_position" property to "DropCapPosition.MARGIN" to place the drop cap
                # outside the left-hand side page margin if our text is left-to-right.
                # Set the "drop_cap_position" property to "DropCapPosition.NORMAL" to place the drop cap within the page margins
                # and to wrap the rest of the text around it.
                # "DropCapPosition.None" is the default state for all paragraphs.
                format = doc.first_section.body.first_paragraph.paragraph_format
                format.drop_cap_position = drop_cap_position

                doc.save(ARTIFACTS_DIR + "ParagraphFormat.drop_cap.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "ParagraphFormat.drop_cap.docx")

                self.assertEqual(drop_cap_position, doc.first_section.body.paragraphs[0].paragraph_format.drop_cap_position)
                self.assertEqual(aw.DropCapPosition.NONE, doc.first_section.body.paragraphs[1].paragraph_format.drop_cap_position)

    def test_line_spacing(self):

        #ExStart
        #ExFor:ParagraphFormat.line_spacing
        #ExFor:ParagraphFormat.line_spacing_rule
        #ExSummary:Shows how to work with line spacing.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Below are three line spacing rules that we can define using the
        # paragraph's "line_spacing_rule" property to configure spacing between paragraphs.
        # 1 -  Set a minimum amount of spacing.
        # This will give vertical padding to lines of text of any size
        # that is too small to maintain the minimum line-height.
        builder.paragraph_format.line_spacing_rule = aw.LineSpacingRule.AT_LEAST
        builder.paragraph_format.line_spacing = 20

        builder.writeln("Minimum line spacing of 20.")
        builder.writeln("Minimum line spacing of 20.")

        # 2 -  Set exact spacing.
        # Using font sizes that are too large for the spacing will truncate the text.
        builder.paragraph_format.line_spacing_rule = aw.LineSpacingRule.EXACTLY
        builder.paragraph_format.line_spacing = 5

        builder.writeln("Line spacing of exactly 5.")
        builder.writeln("Line spacing of exactly 5.")

        # 3 -  Set spacing as a multiple of default line spacing, which is 12 points by default.
        # This kind of spacing will scale to different font sizes.
        builder.paragraph_format.line_spacing_rule = aw.LineSpacingRule.MULTIPLE
        builder.paragraph_format.line_spacing = 18

        builder.writeln("Line spacing of 1.5 default lines.")
        builder.writeln("Line spacing of 1.5 default lines.")

        doc.save(ARTIFACTS_DIR + "ParagraphFormat.line_spacing.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "ParagraphFormat.line_spacing.docx")
        paragraphs = doc.first_section.body.paragraphs

        self.assertEqual(aw.LineSpacingRule.AT_LEAST, paragraphs[0].paragraph_format.line_spacing_rule)
        self.assertEqual(20.0, paragraphs[0].paragraph_format.line_spacing)
        self.assertEqual(aw.LineSpacingRule.AT_LEAST, paragraphs[1].paragraph_format.line_spacing_rule)
        self.assertEqual(20.0, paragraphs[1].paragraph_format.line_spacing)

        self.assertEqual(aw.LineSpacingRule.EXACTLY, paragraphs[2].paragraph_format.line_spacing_rule)
        self.assertEqual(5.0, paragraphs[2].paragraph_format.line_spacing)
        self.assertEqual(aw.LineSpacingRule.EXACTLY, paragraphs[3].paragraph_format.line_spacing_rule)
        self.assertEqual(5.0, paragraphs[3].paragraph_format.line_spacing)

        self.assertEqual(aw.LineSpacingRule.MULTIPLE, paragraphs[4].paragraph_format.line_spacing_rule)
        self.assertEqual(18.0, paragraphs[4].paragraph_format.line_spacing)
        self.assertEqual(aw.LineSpacingRule.MULTIPLE, paragraphs[5].paragraph_format.line_spacing_rule)
        self.assertEqual(18.0, paragraphs[5].paragraph_format.line_spacing)

    def test_paragraph_spacing_auto(self):

        for auto_spacing in (False, True):
            with self.subTest(auto_spacing=auto_spacing):
                #ExStart
                #ExFor:ParagraphFormat.space_after
                #ExFor:ParagraphFormat.space_after_auto
                #ExFor:ParagraphFormat.space_before
                #ExFor:ParagraphFormat.space_before_auto
                #ExSummary:Shows how to set automatic paragraph spacing.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Apply a large amount of spacing before and after paragraphs that this builder will create.
                builder.paragraph_format.space_before = 24
                builder.paragraph_format.space_after = 24

                # Set these flags to "True" to apply automatic spacing,
                # effectively ignoring the spacing in the properties we set above.
                # Leave them as "False" will apply our custom paragraph spacing.
                builder.paragraph_format.space_after_auto = auto_spacing
                builder.paragraph_format.space_before_auto = auto_spacing

                # Insert two paragraphs that will have spacing above and below them and save the document.
                builder.writeln("Paragraph 1.")
                builder.writeln("Paragraph 2.")

                doc.save(ARTIFACTS_DIR + "ParagraphFormat.paragraph_spacing_auto.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "ParagraphFormat.paragraph_spacing_auto.docx")
                format = doc.first_section.body.paragraphs[0].paragraph_format

                self.assertEqual(24.0, format.space_before)
                self.assertEqual(24.0, format.space_after)
                self.assertEqual(auto_spacing, format.space_after_auto)
                self.assertEqual(auto_spacing, format.space_before_auto)

                format = doc.first_section.body.paragraphs[1].paragraph_format

                self.assertEqual(24.0, format.space_before)
                self.assertEqual(24.0, format.space_after)
                self.assertEqual(auto_spacing, format.space_after_auto)
                self.assertEqual(auto_spacing, format.space_before_auto)

    def test_paragraph_spacing_same_style(self):

        for no_space_between_paragraphs_of_same_style in (False, True):
            with self.subTest(no_space_between_paragraphs_of_same_style=no_space_between_paragraphs_of_same_style):
                #ExStart
                #ExFor:ParagraphFormat.space_after
                #ExFor:ParagraphFormat.space_before
                #ExFor:ParagraphFormat.no_space_between_paragraphs_of_same_style
                #ExSummary:Shows how to apply no spacing between paragraphs with the same style.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Apply a large amount of spacing before and after paragraphs that this builder will create.
                builder.paragraph_format.space_before = 24
                builder.paragraph_format.space_after = 24

                # Set the "no_space_between_paragraphs_of_same_style" flag to "True" to apply
                # no spacing between paragraphs with the same style, which will group similar paragraphs.
                # Leave the "no_space_between_paragraphs_of_same_style" flag as "False"
                # to evenly apply spacing to every paragraph.
                builder.paragraph_format.no_space_between_paragraphs_of_same_style = no_space_between_paragraphs_of_same_style

                builder.paragraph_format.style = doc.styles.get_by_name("Normal")
                builder.writeln(f"Paragraph in the \"{builder.paragraph_format.style.name}\" style.")
                builder.writeln(f"Paragraph in the \"{builder.paragraph_format.style.name}\" style.")
                builder.writeln(f"Paragraph in the \"{builder.paragraph_format.style.name}\" style.")
                builder.paragraph_format.style = doc.styles.get_by_name("Quote")
                builder.writeln(f"Paragraph in the \"{builder.paragraph_format.style.name}\" style.")
                builder.writeln(f"Paragraph in the \"{builder.paragraph_format.style.name}\" style.")
                builder.paragraph_format.style = doc.styles.get_by_name("Normal")
                builder.writeln(f"Paragraph in the \"{builder.paragraph_format.style.name}\" style.")
                builder.writeln(f"Paragraph in the \"{builder.paragraph_format.style.name}\" style.")

                doc.save(ARTIFACTS_DIR + "ParagraphFormat.paragraph_spacing_same_style.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "ParagraphFormat.paragraph_spacing_same_style.docx")

                for paragraph in doc.first_section.body.paragraphs:
                    paragraph = paragraph.as_paragraph()
                    format = paragraph.paragraph_format

                    self.assertEqual(24.0, format.space_before)
                    self.assertEqual(24.0, format.space_after)
                    self.assertEqual(no_space_between_paragraphs_of_same_style, format.no_space_between_paragraphs_of_same_style)

    def test_paragraph_outline_level(self):

        #ExStart
        #ExFor:ParagraphFormat.outline_level
        #ExSummary:Shows how to configure paragraph outline levels to create collapsible text.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Each paragraph has an OutlineLevel, which could be any number from 1 to 9, or at the default "BODY_TEXT" value.
        # Setting the property to one of the numbered values will show an arrow to the left
        # of the beginning of the paragraph.
        builder.paragraph_format.outline_level = aw.OutlineLevel.LEVEL1
        builder.writeln("Paragraph outline level 1.")

        # Level 1 is the topmost level. If there is a paragraph with a lower level below a paragraph with a higher level,
        # collapsing the higher-level paragraph will collapse the lower level paragraph.
        builder.paragraph_format.outline_level = aw.OutlineLevel.LEVEL2
        builder.writeln("Paragraph outline level 2.")

        # Two paragraphs of the same level will not collapse each other,
        # and the arrows do not collapse the paragraphs they point to.
        builder.paragraph_format.outline_level = aw.OutlineLevel.LEVEL3
        builder.writeln("Paragraph outline level 3.")
        builder.writeln("Paragraph outline level 3.")

        # The default "BODY_TEXT" value is the lowest, which a paragraph of any level can collapse.
        builder.paragraph_format.outline_level = aw.OutlineLevel.BODY_TEXT
        builder.writeln("Paragraph at main text level.")

        doc.save(ARTIFACTS_DIR + "ParagraphFormat.paragraph_outline_level.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "ParagraphFormat.paragraph_outline_level.docx")
        paragraphs = doc.first_section.body.paragraphs

        self.assertEqual(aw.OutlineLevel.LEVEL1, paragraphs[0].paragraph_format.outline_level)
        self.assertEqual(aw.OutlineLevel.LEVEL2, paragraphs[1].paragraph_format.outline_level)
        self.assertEqual(aw.OutlineLevel.LEVEL3, paragraphs[2].paragraph_format.outline_level)
        self.assertEqual(aw.OutlineLevel.LEVEL3, paragraphs[3].paragraph_format.outline_level)
        self.assertEqual(aw.OutlineLevel.BODY_TEXT, paragraphs[4].paragraph_format.outline_level)

    def test_page_break_before(self):

        for page_break_before in (False, True):
            with self.subTest(page_break_before=page_break_before):
                #ExStart
                #ExFor:ParagraphFormat.page_break_before
                #ExSummary:Shows how to create paragraphs with page breaks at the beginning.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # Set this flag to "True" to apply a page break to each paragraph's beginning
                # that the document builder will create under this ParagraphFormat configuration.
                # The first paragraph will not receive a page break.
                # Leave this flag as "False" to start each new paragraph on the same page
                # as the previous, provided there is sufficient space.
                builder.paragraph_format.page_break_before = page_break_before

                builder.writeln("Paragraph 1.")
                builder.writeln("Paragraph 2.")

                layout_collector = aw.layout.LayoutCollector(doc)
                paragraphs = doc.first_section.body.paragraphs

                if page_break_before:
                    self.assertEqual(1, layout_collector.get_start_page_index(paragraphs[0]))
                    self.assertEqual(2, layout_collector.get_start_page_index(paragraphs[1]))
                else:
                    self.assertEqual(1, layout_collector.get_start_page_index(paragraphs[0]))
                    self.assertEqual(1, layout_collector.get_start_page_index(paragraphs[1]))

                doc.save(ARTIFACTS_DIR + "ParagraphFormat.page_break_before.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "ParagraphFormat.page_break_before.docx")
                paragraphs = doc.first_section.body.paragraphs

                self.assertEqual(page_break_before, paragraphs[0].paragraph_format.page_break_before)
                self.assertEqual(page_break_before, paragraphs[1].paragraph_format.page_break_before)

    def test_widow_control(self):

        for widow_control in (False, True):
            with self.subTest(widow_control=widow_control):
                #ExStart
                #ExFor:ParagraphFormat.widow_control
                #ExSummary:Shows how to enable widow/orphan control for a paragraph.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                # When we write the text that does not fit onto one page, one line may spill over onto the next page.
                # The single line that ends up on the next page is called an "Orphan",
                # and the previous line where the orphan broke off is called a "Widow".
                # We can fix orphans and widows by rearranging text via font size, spacing, or page margins.
                # If we wish to preserve our document's dimensions, we can set this flag to "True"
                # to push widows onto the same page as their respective orphans.
                # Leave this flag as "False" will leave widow/orphan pairs in text.
                # Every paragraph has this setting accessible in Microsoft Word via Home -> Paragraph -> Paragraph Settings
                # (button on bottom right hand corner of "Paragraph" tab) -> "Widow/Orphan control".
                builder.paragraph_format.widow_control = widow_control

                # Insert text that produces an orphan and a widow.
                builder.font.size = 68
                builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                                "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")

                doc.save(ARTIFACTS_DIR + "ParagraphFormat.widow_control.docx")
                #ExEnd

                doc = aw.Document(ARTIFACTS_DIR + "ParagraphFormat.widow_control.docx")

                self.assertEqual(widow_control, doc.first_section.body.paragraphs[0].paragraph_format.widow_control)

    def test_lines_to_drop(self):

        #ExStart
        #ExFor:ParagraphFormat.lines_to_drop
        #ExSummary:Shows how to set the size of a drop cap.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Modify the "lines_to_drop" property to designate a paragraph as a drop cap,
        # which will turn it into a large capital letter that will decorate the next paragraph.
        # Give this property a value of 4 to give the drop cap the height of four text lines.
        builder.paragraph_format.lines_to_drop = 4
        builder.writeln("H")

        # Reset the "lines_to_drop" property to 0 to turn the next paragraph into an ordinary paragraph.
        # The text in this paragraph will wrap around the drop cap.
        builder.paragraph_format.lines_to_drop = 0
        builder.writeln("ello world!")

        doc.save(ARTIFACTS_DIR + "ParagraphFormat.lines_to_drop.odt")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "ParagraphFormat.lines_to_drop.odt")
        paragraphs = doc.first_section.body.paragraphs

        self.assertEqual(4, paragraphs[0].paragraph_format.lines_to_drop)
        self.assertEqual(0, paragraphs[1].paragraph_format.lines_to_drop)

    def test_suppress_hyphens(self):

        for suppress_auto_hyphens in (False, True):
            with self.subTest(suppress_auto_hyphens=suppress_auto_hyphens):
                #ExStart
                #ExFor:ParagraphFormat.suppress_auto_hyphens
                #ExSummary:Shows how to suppress hyphenation for a paragraph.
                aw.Hyphenation.register_dictionary("de-CH", MY_DIR + "hyph_de_CH.dic")

                self.assertTrue(aw.Hyphenation.is_dictionary_registered("de-CH"))

                # Open a document containing text with a locale matching that of our dictionary.
                # When we save this document to a fixed page save format, its text will have hyphenation.
                doc = aw.Document(MY_DIR + "German text.docx")

                # We can set the "suppress_auto_hyphens" property to "True" to disable hyphenation
                # for a specific paragraph while keeping it enabled for the rest of the document.
                # The default value for this property is "False",
                # which means every paragraph by default uses hyphenation if any is available.
                doc.first_section.body.first_paragraph.paragraph_format.suppress_auto_hyphens = suppress_auto_hyphens

                doc.save(ARTIFACTS_DIR + "ParagraphFormat.suppress_hyphens.pdf")
                #ExEnd

                #pdf_doc = aspose.pdf.Document(ARTIFACTS_DIR + "ParagraphFormat.suppress_hyphens.pdf")
                #text_absorber = aspose.pdf.text.TextAbsorber()
                #textAbsorber.visit(pdf_doc)

                #if suppress_auto_hyphens:
                #    self.assertTrue(textAbsorber.text.contains(
                #        "La  ob  storen  an  deinen  am  sachen. \r\n" +
                #        "Doppelte  um  da  am  spateren  verlogen \r\n" +
                #        "gekommen  achtzehn  blaulich."))
                #else:
                #    self.assertTrue(textAbsorber.text.contains(
                #        "La ob storen an deinen am sachen. Dop-\r\n" +
                #        "pelte  um  da  am  spateren  verlogen  ge-\r\n" +
                #        "kommen  achtzehn  blaulich."))

    def test_paragraph_spacing_and_indents(self):

        #ExStart
        #ExFor:ParagraphFormat.character_unit_left_indent
        #ExFor:ParagraphFormat.character_unit_right_indent
        #ExFor:ParagraphFormat.character_unit_first_line_indent
        #ExFor:ParagraphFormat.line_unit_before
        #ExFor:ParagraphFormat.line_unit_after
        #ExSummary:Shows how to change paragraph spacing and indents.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        format = doc.first_section.body.first_paragraph.paragraph_format

        # Below are five different spacing options, along with the properties that their configuration indirectly affects.
        # 1 -  Left indent:
        self.assertEqual(format.left_indent, 0.0)

        format.character_unit_left_indent = 10.0

        self.assertEqual(format.left_indent, 120.0)

        # 2 -  Right indent:
        self.assertEqual(format.right_indent, 0.0)

        format.character_unit_right_indent = -5.5

        self.assertEqual(format.right_indent, -66.0)

        # 3 -  Hanging indent:
        self.assertEqual(format.first_line_indent, 0.0)

        format.character_unit_first_line_indent = 20.3

        self.assertAlmostEqual(format.first_line_indent, 243.59, delta=0.1)

        # 4 -  Line spacing before paragraphs:
        self.assertEqual(format.space_before, 0.0)

        format.line_unit_before = 5.1

        self.assertAlmostEqual(format.space_before, 61.1, delta=0.1)

        # 5 -  Line spacing after paragraphs:
        self.assertEqual(format.space_after, 0.0)

        format.line_unit_after = 10.9

        self.assertAlmostEqual(format.space_after, 130.8, delta=0.1)

        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                        "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.")
        builder.write("测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试" +
                      "文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档测试文档")
        #ExEnd

        doc = DocumentHelper.save_open(doc)
        format = doc.first_section.body.first_paragraph.paragraph_format

        self.assertEqual(format.character_unit_left_indent, 10.0)
        self.assertEqual(format.left_indent, 120.0)

        self.assertEqual(format.character_unit_right_indent, -5.5)
        self.assertEqual(format.right_indent, -66.0)

        self.assertEqual(format.character_unit_first_line_indent, 20.3)
        self.assertAlmostEqual(format.first_line_indent, 243.59, delta=0.1)

        self.assertAlmostEqual(format.line_unit_before, 5.1, delta=0.1)
        self.assertAlmostEqual(format.space_before, 61.1, delta=0.1)

        self.assertEqual(format.line_unit_after, 10.9)
        self.assertAlmostEqual(format.space_after, 130.8, delta=0.1)
