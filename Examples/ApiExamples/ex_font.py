# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import sys
import pathlib
import glob
import aspose.pydrawing
import aspose.words as aw
import aspose.words.fonts
import aspose.words.settings
import aspose.words.themes
import os
import system_helper
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, FONTS_DIR, MY_DIR

class ExFont(ApiExampleBase):

    def test_create_formatted_run(self):
        #ExStart
        #ExFor:Document.__init__
        #ExFor:Font
        #ExFor:Font.name
        #ExFor:Font.size
        #ExFor:Font.highlight_color
        #ExFor:Run
        #ExFor:Run.__init__(DocumentBase,str)
        #ExFor:Story.first_paragraph
        #ExSummary:Shows how to format a run of text using its font property.
        doc = aw.Document()
        run = aw.Run(doc=doc, text='Hello world!')
        font = run.font
        font.name = 'Courier New'
        font.size = 36
        font.highlight_color = aspose.pydrawing.Color.yellow
        doc.first_section.body.first_paragraph.append_child(run)
        doc.save(file_name=ARTIFACTS_DIR + 'Font.CreateFormattedRun.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.CreateFormattedRun.docx')
        run = doc.first_section.body.first_paragraph.runs[0]
        self.assertEqual('Hello world!', run.get_text().strip())
        self.assertEqual('Courier New', run.font.name)
        self.assertEqual(36, run.font.size)
        self.assertEqual(aspose.pydrawing.Color.yellow.to_argb(), run.font.highlight_color.to_argb())

    def test_caps(self):
        #ExStart
        #ExFor:Font.all_caps
        #ExFor:Font.small_caps
        #ExSummary:Shows how to format a run to display its contents in capitals.
        doc = aw.Document()
        para = doc.get_child(aw.NodeType.PARAGRAPH, 0, True).as_paragraph()
        # There are two ways of getting a run to display its lowercase text in uppercase without changing the contents.
        # 1 -  Set the AllCaps flag to display all characters in regular capitals:
        run = aw.Run(doc=doc, text='all capitals')
        run.font.all_caps = True
        para.append_child(run)
        para = para.parent_node.append_child(aw.Paragraph(doc)).as_paragraph()
        # 2 -  Set the SmallCaps flag to display all characters in small capitals:
        # If a character is lower case, it will appear in its upper case form
        # but will have the same height as the lower case (the font's x-height).
        # Characters that were in upper case originally will look the same.
        run = aw.Run(doc=doc, text='Small Capitals')
        run.font.small_caps = True
        para.append_child(run)
        doc.save(file_name=ARTIFACTS_DIR + 'Font.Caps.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.Caps.docx')
        run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('all capitals', run.get_text().strip())
        self.assertTrue(run.font.all_caps)
        run = doc.first_section.body.paragraphs[1].runs[0]
        self.assertEqual('Small Capitals', run.get_text().strip())
        self.assertTrue(run.font.small_caps)

    def test_default_values_embedded_fonts_parameters(self):
        doc = aw.Document()
        self.assertFalse(doc.font_infos.embed_true_type_fonts)
        self.assertFalse(doc.font_infos.embed_system_fonts)
        self.assertFalse(doc.font_infos.save_subset_fonts)

    @unittest.skipUnless(sys.platform.startswith('win'), 'requires Windows')
    def test_font_info_collection(self):
        for embed_all_fonts in [False, True]:
            #ExStart
            #ExFor:FontInfoCollection
            #ExFor:DocumentBase.font_infos
            #ExFor:FontInfoCollection.embed_true_type_fonts
            #ExFor:FontInfoCollection.embed_system_fonts
            #ExFor:FontInfoCollection.save_subset_fonts
            #ExSummary:Shows how to save a document with embedded TrueType fonts.
            doc = aw.Document(file_name=MY_DIR + 'Document.docx')
            font_infos = doc.font_infos
            font_infos.embed_true_type_fonts = embed_all_fonts
            font_infos.embed_system_fonts = embed_all_fonts
            font_infos.save_subset_fonts = embed_all_fonts
            doc.save(file_name=ARTIFACTS_DIR + 'Font.FontInfoCollection.docx')
            #ExEnd
            tested_file_length = system_helper.io.FileInfo(ARTIFACTS_DIR + 'Font.FontInfoCollection.docx').length()
            if embed_all_fonts:
                self.assertTrue(tested_file_length < 28000)
            else:
                self.assertTrue(tested_file_length < 13000)

    def test_work_with_embedded_fonts(self):
        for embed_true_type_fonts, embed_system_fonts, save_subset_fonts in [(True, False, False), (True, True, False), (True, True, True), (True, False, True), (False, False, False)]:
            doc = aw.Document(file_name=MY_DIR + 'Document.docx')
            font_infos = doc.font_infos
            font_infos.embed_true_type_fonts = embed_true_type_fonts
            font_infos.embed_system_fonts = embed_system_fonts
            font_infos.save_subset_fonts = save_subset_fonts
            doc.save(file_name=ARTIFACTS_DIR + 'Font.WorkWithEmbeddedFonts.docx')

    def test_strike_through(self):
        #ExStart
        #ExFor:Font.strike_through
        #ExFor:Font.double_strike_through
        #ExSummary:Shows how to add a line strikethrough to text.
        doc = aw.Document()
        para = doc.get_child(aw.NodeType.PARAGRAPH, 0, True).as_paragraph()
        run = aw.Run(doc=doc, text='Text with a single-line strikethrough.')
        run.font.strike_through = True
        para.append_child(run)
        para = para.parent_node.append_child(aw.Paragraph(doc)).as_paragraph()
        run = aw.Run(doc=doc, text='Text with a double-line strikethrough.')
        run.font.double_strike_through = True
        para.append_child(run)
        doc.save(file_name=ARTIFACTS_DIR + 'Font.StrikeThrough.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.StrikeThrough.docx')
        run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('Text with a single-line strikethrough.', run.get_text().strip())
        self.assertTrue(run.font.strike_through)
        run = doc.first_section.body.paragraphs[1].runs[0]
        self.assertEqual('Text with a double-line strikethrough.', run.get_text().strip())
        self.assertTrue(run.font.double_strike_through)

    def test_position_subscript(self):
        #ExStart
        #ExFor:Font.position
        #ExFor:Font.subscript
        #ExFor:Font.superscript
        #ExSummary:Shows how to format text to offset its position.
        doc = aw.Document()
        para = doc.get_child(aw.NodeType.PARAGRAPH, 0, True).as_paragraph()
        # Raise this run of text 5 points above the baseline.
        run = aw.Run(doc=doc, text='Raised text. ')
        run.font.position = 5
        para.append_child(run)
        # Lower this run of text 10 points below the baseline.
        run = aw.Run(doc=doc, text='Lowered text. ')
        run.font.position = -10
        para.append_child(run)
        # Add a run of normal text.
        run = aw.Run(doc=doc, text='Text in its default position. ')
        para.append_child(run)
        # Add a run of text that appears as subscript.
        run = aw.Run(doc=doc, text='Subscript. ')
        run.font.subscript = True
        para.append_child(run)
        # Add a run of text that appears as superscript.
        run = aw.Run(doc=doc, text='Superscript.')
        run.font.superscript = True
        para.append_child(run)
        doc.save(file_name=ARTIFACTS_DIR + 'Font.PositionSubscript.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.PositionSubscript.docx')
        run = doc.first_section.body.first_paragraph.runs[0]
        self.assertEqual('Raised text.', run.get_text().strip())
        self.assertEqual(5, run.font.position)
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.PositionSubscript.docx')
        run = doc.first_section.body.first_paragraph.runs[1]
        self.assertEqual('Lowered text.', run.get_text().strip())
        self.assertEqual(-10, run.font.position)
        run = doc.first_section.body.first_paragraph.runs[3]
        self.assertEqual('Subscript.', run.get_text().strip())
        self.assertTrue(run.font.subscript)
        run = doc.first_section.body.first_paragraph.runs[4]
        self.assertEqual('Superscript.', run.get_text().strip())
        self.assertTrue(run.font.superscript)

    def test_scaling_spacing(self):
        #ExStart
        #ExFor:Font.scaling
        #ExFor:Font.spacing
        #ExSummary:Shows how to set horizontal scaling and spacing for characters.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Add run of text and increase character width to 150%.
        builder.font.scaling = 150
        builder.writeln('Wide characters')
        # Add run of text and add 1pt of extra horizontal spacing between each character.
        builder.font.spacing = 1
        builder.writeln('Expanded by 1pt')
        # Add run of text and bring characters closer together by 1pt.
        builder.font.spacing = -1
        builder.writeln('Condensed by 1pt')
        doc.save(file_name=ARTIFACTS_DIR + 'Font.ScalingSpacing.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.ScalingSpacing.docx')
        run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('Wide characters', run.get_text().strip())
        self.assertEqual(150, run.font.scaling)
        run = doc.first_section.body.paragraphs[1].runs[0]
        self.assertEqual('Expanded by 1pt', run.get_text().strip())
        self.assertEqual(1, run.font.spacing)
        run = doc.first_section.body.paragraphs[2].runs[0]
        self.assertEqual('Condensed by 1pt', run.get_text().strip())
        self.assertEqual(-1, run.font.spacing)

    def test_italic(self):
        #ExStart
        #ExFor:Font.italic
        #ExSummary:Shows how to write italicized text using a document builder.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.size = 36
        builder.font.italic = True
        builder.writeln('Hello world!')
        doc.save(file_name=ARTIFACTS_DIR + 'Font.Italic.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.Italic.docx')
        run = doc.first_section.body.first_paragraph.runs[0]
        self.assertEqual('Hello world!', run.get_text().strip())
        self.assertTrue(run.font.italic)

    def test_engrave_emboss(self):
        #ExStart
        #ExFor:Font.emboss
        #ExFor:Font.engrave
        #ExSummary:Shows how to apply engraving/embossing effects to text.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.size = 36
        builder.font.color = aspose.pydrawing.Color.light_blue
        # Below are two ways of using shadows to apply a 3D-like effect to the text.
        # 1 -  Engrave text to make it look like the letters are sunken into the page:
        builder.font.engrave = True
        builder.writeln('This text is engraved.')
        # 2 -  Emboss text to make it look like the letters pop out of the page:
        builder.font.engrave = False
        builder.font.emboss = True
        builder.writeln('This text is embossed.')
        doc.save(file_name=ARTIFACTS_DIR + 'Font.EngraveEmboss.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.EngraveEmboss.docx')
        run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('This text is engraved.', run.get_text().strip())
        self.assertTrue(run.font.engrave)
        self.assertFalse(run.font.emboss)
        run = doc.first_section.body.paragraphs[1].runs[0]
        self.assertEqual('This text is embossed.', run.get_text().strip())
        self.assertFalse(run.font.engrave)
        self.assertTrue(run.font.emboss)

    def test_shadow(self):
        #ExStart
        #ExFor:Font.shadow
        #ExSummary:Shows how to create a run of text formatted with a shadow.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Set the Shadow flag to apply an offset shadow effect,
        # making it look like the letters are floating above the page.
        builder.font.shadow = True
        builder.font.size = 36
        builder.writeln('This text has a shadow.')
        doc.save(file_name=ARTIFACTS_DIR + 'Font.Shadow.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.Shadow.docx')
        run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('This text has a shadow.', run.get_text().strip())
        self.assertTrue(run.font.shadow)

    def test_outline(self):
        #ExStart
        #ExFor:Font.outline
        #ExSummary:Shows how to create a run of text formatted as outline.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Set the Outline flag to change the text's fill color to white and
        # leave a thin outline around each character in the original color of the text.
        builder.font.outline = True
        builder.font.color = aspose.pydrawing.Color.blue
        builder.font.size = 36
        builder.writeln('This text has an outline.')
        doc.save(file_name=ARTIFACTS_DIR + 'Font.Outline.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.Outline.docx')
        run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('This text has an outline.', run.get_text().strip())
        self.assertTrue(run.font.outline)

    def test_hidden(self):
        #ExStart
        #ExFor:Font.hidden
        #ExSummary:Shows how to create a run of hidden text.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # With the Hidden flag set to true, any text that we create using this Font object will be invisible in the document.
        # We will not see or highlight hidden text unless we enable the "Hidden text" option
        # found in Microsoft Word via "File" -> "Options" -> "Display". The text will still be there,
        # and we will be able to access this text programmatically.
        # It is not advised to use this method to hide sensitive information.
        builder.font.hidden = True
        builder.font.size = 36
        builder.writeln('This text will not be visible in the document.')
        doc.save(file_name=ARTIFACTS_DIR + 'Font.Hidden.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.Hidden.docx')
        run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('This text will not be visible in the document.', run.get_text().strip())
        self.assertTrue(run.font.hidden)

    def test_kerning(self):
        #ExStart
        #ExFor:Font.kerning
        #ExSummary:Shows how to specify the font size at which kerning begins to take effect.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.name = 'Arial Black'
        # Set the builder's font size, and minimum size at which kerning will take effect.
        # The font size falls below the kerning threshold, so the run bellow will not have kerning.
        builder.font.size = 18
        builder.font.kerning = 24
        builder.writeln('TALLY. (Kerning not applied)')
        # Set the kerning threshold so that the builder's current font size is above it.
        # Any text we add from this point will have kerning applied. The spaces between characters
        # will be adjusted, normally resulting in a slightly more aesthetically pleasing text run.
        builder.font.kerning = 12
        builder.writeln('TALLY. (Kerning applied)')
        doc.save(file_name=ARTIFACTS_DIR + 'Font.Kerning.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.Kerning.docx')
        run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('TALLY. (Kerning not applied)', run.get_text().strip())
        self.assertEqual(24, run.font.kerning)
        self.assertEqual(18, run.font.size)
        run = doc.first_section.body.paragraphs[1].runs[0]
        self.assertEqual('TALLY. (Kerning applied)', run.get_text().strip())
        self.assertEqual(12, run.font.kerning)
        self.assertEqual(18, run.font.size)

    def test_no_proofing(self):
        #ExStart
        #ExFor:Font.no_proofing
        #ExSummary:Shows how to prevent text from being spell checked by Microsoft Word.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Normally, Microsoft Word emphasizes spelling errors with a jagged red underline.
        # We can un-set the "NoProofing" flag to create a portion of text that
        # bypasses the spell checker while completely disabling it.
        builder.font.no_proofing = True
        builder.writeln('Proofing has been disabled, so these spelking errrs will not display red lines underneath.')
        doc.save(file_name=ARTIFACTS_DIR + 'Font.NoProofing.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.NoProofing.docx')
        run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('Proofing has been disabled, so these spelking errrs will not display red lines underneath.', run.get_text().strip())
        self.assertTrue(run.font.no_proofing)

    def test_underlines(self):
        #ExStart
        #ExFor:Font.underline
        #ExFor:Font.underline_color
        #ExSummary:Shows how to configure the style and color of a text underline.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.underline = aw.Underline.DOTTED
        builder.font.underline_color = aspose.pydrawing.Color.red
        builder.writeln('Underlined text.')
        doc.save(file_name=ARTIFACTS_DIR + 'Font.Underlines.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.Underlines.docx')
        run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('Underlined text.', run.get_text().strip())
        self.assertEqual(aw.Underline.DOTTED, run.font.underline)
        self.assertEqual(aspose.pydrawing.Color.red.to_argb(), run.font.underline_color.to_argb())

    def test_complex_script(self):
        #ExStart
        #ExFor:Font.complex_script
        #ExSummary:Shows how to add text that is always treated as complex script.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.complex_script = True
        builder.writeln('Text treated as complex script.')
        doc.save(file_name=ARTIFACTS_DIR + 'Font.ComplexScript.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.ComplexScript.docx')
        run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('Text treated as complex script.', run.get_text().strip())
        self.assertTrue(run.font.complex_script)

    def test_sparkling_text(self):
        #ExStart
        #ExFor:TextEffect
        #ExFor:Font.text_effect
        #ExSummary:Shows how to apply a visual effect to a run.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.size = 36
        builder.font.text_effect = aw.TextEffect.SPARKLE_TEXT
        builder.writeln('Text with a sparkle effect.')
        # Older versions of Microsoft Word only support font animation effects.
        doc.save(file_name=ARTIFACTS_DIR + 'Font.SparklingText.doc')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.SparklingText.doc')
        run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('Text with a sparkle effect.', run.get_text().strip())
        self.assertEqual(aw.TextEffect.SPARKLE_TEXT, run.font.text_effect)

    def test_foreground_and_background(self):
        #ExStart
        #ExFor:Shading.foreground_pattern_theme_color
        #ExFor:Shading.background_pattern_theme_color
        #ExFor:Shading.foreground_tint_and_shade
        #ExFor:Shading.background_tint_and_shade
        #ExSummary:Shows how to set foreground and background colors for shading texture.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        shading = doc.first_section.body.first_paragraph.paragraph_format.shading
        shading.texture = aw.TextureIndex.TEXTURE_12PT5_PERCENT
        shading.foreground_pattern_theme_color = aw.themes.ThemeColor.DARK1
        shading.background_pattern_theme_color = aw.themes.ThemeColor.DARK2
        shading.foreground_tint_and_shade = 0.5
        shading.background_tint_and_shade = -0.2
        builder.font.border.color = aspose.pydrawing.Color.green
        builder.font.border.line_width = 2.5
        builder.font.border.line_style = aw.LineStyle.DASH_DOT_STROKER
        builder.writeln('Foreground and background pattern colors for shading texture.')
        doc.save(file_name=ARTIFACTS_DIR + 'Font.ForegroundAndBackground.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.ForegroundAndBackground.docx')
        run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('Foreground and background pattern colors for shading texture.', run.get_text().strip())
        self.assertEqual(aw.themes.ThemeColor.DARK1, doc.first_section.body.paragraphs[0].paragraph_format.shading.foreground_pattern_theme_color)
        self.assertEqual(aw.themes.ThemeColor.DARK2, doc.first_section.body.paragraphs[0].paragraph_format.shading.background_pattern_theme_color)
        self.assertAlmostEqual(0.5, doc.first_section.body.paragraphs[0].paragraph_format.shading.foreground_tint_and_shade, delta=0.1)
        self.assertAlmostEqual(-0.2, doc.first_section.body.paragraphs[0].paragraph_format.shading.background_tint_and_shade, delta=0.1)

    def test_shading(self):
        #ExStart
        #ExFor:Font.shading
        #ExSummary:Shows how to apply shading to text created by a document builder.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.color = aspose.pydrawing.Color.white
        # One way to make the text created using our white font color visible
        # is to apply a background shading effect.
        shading = builder.font.shading
        shading.texture = aw.TextureIndex.TEXTURE_DIAGONAL_UP
        shading.background_pattern_color = aspose.pydrawing.Color.orange_red
        shading.foreground_pattern_color = aspose.pydrawing.Color.dark_blue
        builder.writeln('White text on an orange background with a two-tone texture.')
        doc.save(file_name=ARTIFACTS_DIR + 'Font.Shading.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.Shading.docx')
        run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('White text on an orange background with a two-tone texture.', run.get_text().strip())
        self.assertEqual(aspose.pydrawing.Color.white.to_argb(), run.font.color.to_argb())
        self.assertEqual(aw.TextureIndex.TEXTURE_DIAGONAL_UP, run.font.shading.texture)
        self.assertEqual(aspose.pydrawing.Color.orange_red.to_argb(), run.font.shading.background_pattern_color.to_argb())
        self.assertEqual(aspose.pydrawing.Color.dark_blue.to_argb(), run.font.shading.foreground_pattern_color.to_argb())

    def test_name_ascii(self):
        #ExStart
        #ExFor:Font.name_ascii
        #ExFor:Font.name_other
        #ExSummary:Shows how Microsoft Word can combine two different fonts in one run.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Suppose a run that we use the builder to insert while using this font configuration
        # contains characters within the ASCII characters' range. In that case,
        # it will display those characters using this font.
        builder.font.name_ascii = 'Calibri'
        # With no other font specified, the builder will also apply this font to all characters that it inserts.
        self.assertEqual('Calibri', builder.font.name)
        # Specify a font to use for all characters outside of the ASCII range.
        # Ideally, this font should have a glyph for each required non-ASCII character code.
        builder.font.name_other = 'Courier New'
        # Insert a run with one word consisting of ASCII characters, and one word with all characters outside that range.
        # Each character will be displayed using either of the fonts, depending on.
        builder.writeln('Hello, Привет')
        doc.save(file_name=ARTIFACTS_DIR + 'Font.NameAscii.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.NameAscii.docx')
        run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('Hello, Привет', run.get_text().strip())
        self.assertEqual('Calibri', run.font.name)
        self.assertEqual('Calibri', run.font.name_ascii)
        self.assertEqual('Courier New', run.font.name_other)

    def test_change_style(self):
        #ExStart
        #ExFor:Font.style_name
        #ExFor:Font.style_identifier
        #ExFor:StyleIdentifier
        #ExSummary:Shows how to change the style of existing text.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Below are two ways of referencing styles.
        # 1 -  Using the style name:
        builder.font.style_name = 'Emphasis'
        builder.writeln('Text originally in "Emphasis" style')
        # 2 -  Using a built-in style identifier:
        builder.font.style_identifier = aw.StyleIdentifier.INTENSE_EMPHASIS
        builder.writeln('Text originally in "Intense Emphasis" style')
        # Convert all uses of one style to another,
        # using the above methods to reference old and new styles.
        for run in doc.get_child_nodes(aw.NodeType.RUN, True):
            run = run.as_run()
            if run.font.style_name == 'Emphasis':
                run.font.style_name = 'Strong'
            if run.font.style_identifier == aw.StyleIdentifier.INTENSE_EMPHASIS:
                run.font.style_identifier = aw.StyleIdentifier.STRONG
        doc.save(file_name=ARTIFACTS_DIR + 'Font.ChangeStyle.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.ChangeStyle.docx')
        doc_run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('Text originally in "Emphasis" style', doc_run.get_text().strip())
        self.assertEqual(aw.StyleIdentifier.STRONG, doc_run.font.style_identifier)
        self.assertEqual('Strong', doc_run.font.style_name)
        doc_run = doc.first_section.body.paragraphs[1].runs[0]
        self.assertEqual('Text originally in "Intense Emphasis" style', doc_run.get_text().strip())
        self.assertEqual(aw.StyleIdentifier.STRONG, doc_run.font.style_identifier)
        self.assertEqual('Strong', doc_run.font.style_name)

    def test_built_in(self):
        #ExStart
        #ExFor:Style.built_in
        #ExSummary:Shows how to differentiate custom styles from built-in styles.
        doc = aw.Document()
        # When we create a document using Microsoft Word, or programmatically using Aspose.Words,
        # the document will come with a collection of styles to apply to its text to modify its appearance.
        # We can access these built-in styles via the document's "Styles" collection.
        # These styles will all have the "BuiltIn" flag set to "true".
        style = doc.styles.get_by_name('Emphasis')
        self.assertTrue(style.built_in)
        # Create a custom style and add it to the collection.
        # Custom styles such as this will have the "BuiltIn" flag set to "false".
        style = doc.styles.add(aw.StyleType.CHARACTER, 'MyStyle')
        style.font.color = aspose.pydrawing.Color.navy
        style.font.name = 'Courier New'
        self.assertFalse(style.built_in)
        #ExEnd

    def test_style(self):
        #ExStart
        #ExFor:Font.style
        #ExSummary:Applies a double underline to all runs in a document that are formatted with custom character styles.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert a custom style and apply it to text created using a document builder.
        style = doc.styles.add(aw.StyleType.CHARACTER, 'MyStyle')
        style.font.color = aspose.pydrawing.Color.red
        style.font.name = 'Courier New'
        builder.font.style_name = 'MyStyle'
        builder.write('This text is in a custom style.')
        # Iterate over every run and add a double underline to every custom style.
        for run in doc.get_child_nodes(aw.NodeType.RUN, True):
            run = run.as_run()
            char_style = run.font.style
            if not char_style.built_in:
                run.font.underline = aw.Underline.DOUBLE
        doc.save(file_name=ARTIFACTS_DIR + 'Font.Style.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.Style.docx')
        doc_run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('This text is in a custom style.', doc_run.get_text().strip())
        self.assertEqual('MyStyle', doc_run.font.style_name)
        self.assertFalse(doc_run.font.style.built_in)
        self.assertEqual(aw.Underline.DOUBLE, doc_run.font.underline)

    def test_set_font_auto_color(self):
        #ExStart
        #ExFor:Font.auto_color
        #ExSummary:Shows how to improve readability by automatically selecting text color based on the brightness of its background.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # If a run's Font object does not specify text color, it will automatically
        # select either black or white depending on the background color's color.
        self.assertEqual(aspose.pydrawing.Color.empty().to_argb(), builder.font.color.to_argb())
        # The default color for text is black. If the color of the background is dark, black text will be difficult to see.
        # To solve this problem, the AutoColor property will display this text in white.
        builder.font.shading.background_pattern_color = aspose.pydrawing.Color.dark_blue
        builder.writeln('The text color automatically chosen for this run is white.')
        self.assertEqual(aspose.pydrawing.Color.white.to_argb(), doc.first_section.body.paragraphs[0].runs[0].font.auto_color.to_argb())
        # If we change the background to a light color, black will be a more
        # suitable text color than white so that the auto color will display it in black.
        builder.font.shading.background_pattern_color = aspose.pydrawing.Color.light_blue
        builder.writeln('The text color automatically chosen for this run is black.')
        self.assertEqual(aspose.pydrawing.Color.black.to_argb(), doc.first_section.body.paragraphs[1].runs[0].font.auto_color.to_argb())
        doc.save(file_name=ARTIFACTS_DIR + 'Font.SetFontAutoColor.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Font.SetFontAutoColor.docx')
        run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('The text color automatically chosen for this run is white.', run.get_text().strip())
        self.assertEqual(aspose.pydrawing.Color.empty().to_argb(), run.font.color.to_argb())
        self.assertEqual(aspose.pydrawing.Color.dark_blue.to_argb(), run.font.shading.background_pattern_color.to_argb())
        run = doc.first_section.body.paragraphs[1].runs[0]
        self.assertEqual('The text color automatically chosen for this run is black.', run.get_text().strip())
        self.assertEqual(aspose.pydrawing.Color.empty().to_argb(), run.font.color.to_argb())
        self.assertEqual(aspose.pydrawing.Color.light_blue.to_argb(), run.font.shading.background_pattern_color.to_argb())

    def test_default_fonts(self):
        #ExStart
        #ExFor:FontInfoCollection.contains(str)
        #ExFor:FontInfoCollection.count
        #ExSummary:Shows info about the fonts that are present in the blank document.
        doc = aw.Document()
        # A blank document contains 3 default fonts. Each font in the document
        # will have a corresponding FontInfo object which contains details about that font.
        self.assertEqual(3, doc.font_infos.count)
        self.assertTrue(doc.font_infos.contains('Times New Roman'))
        self.assertEqual(204, doc.font_infos.get_by_name('Times New Roman').charset)
        self.assertTrue(doc.font_infos.contains('Symbol'))
        self.assertTrue(doc.font_infos.contains('Arial'))
        #ExEnd

    def test_extract_embedded_font(self):
        #ExStart
        #ExFor:EmbeddedFontFormat
        #ExFor:EmbeddedFontStyle
        #ExFor:FontInfo.get_embedded_font(EmbeddedFontFormat,EmbeddedFontStyle)
        #ExFor:FontInfo.get_embedded_font_as_open_type(EmbeddedFontStyle)
        #ExFor:FontInfoCollection.__getitem__(int)
        #ExFor:FontInfoCollection.__getitem__(str)
        #ExSummary:Shows how to extract an embedded font from a document, and save it to the local file system.
        doc = aw.Document(file_name=MY_DIR + 'Embedded font.docx')
        embedded_font = doc.font_infos.get_by_name('Alte DIN 1451 Mittelschrift')
        embedded_font_bytes = embedded_font.get_embedded_font(aw.fonts.EmbeddedFontFormat.OPEN_TYPE, aw.fonts.EmbeddedFontStyle.REGULAR)
        self.assertIsNotNone(embedded_font_bytes)  #ExSkip
        system_helper.io.File.write_all_bytes(ARTIFACTS_DIR + 'Alte DIN 1451 Mittelschrift.ttf', embedded_font_bytes)
        # Embedded font formats may be different in other formats such as .doc.
        # We need to know the correct format before we can extract the font.
        doc = aw.Document(file_name=MY_DIR + 'Embedded font.doc')
        self.assertIsNone(doc.font_infos.get_by_name('Alte DIN 1451 Mittelschrift').get_embedded_font(aw.fonts.EmbeddedFontFormat.OPEN_TYPE, aw.fonts.EmbeddedFontStyle.REGULAR))
        self.assertIsNotNone(doc.font_infos.get_by_name('Alte DIN 1451 Mittelschrift').get_embedded_font(aw.fonts.EmbeddedFontFormat.EMBEDDED_OPEN_TYPE, aw.fonts.EmbeddedFontStyle.REGULAR))
        # Also, we can convert embedded OpenType format, which comes from .doc documents, to OpenType.
        embedded_font_bytes = doc.font_infos.get_by_name('Alte DIN 1451 Mittelschrift').get_embedded_font_as_open_type(aw.fonts.EmbeddedFontStyle.REGULAR)
        system_helper.io.File.write_all_bytes(ARTIFACTS_DIR + 'Alte DIN 1451 Mittelschrift.otf', embedded_font_bytes)
        #ExEnd

    @unittest.skipUnless(sys.platform.startswith('win'), 'different calculation on Linux')
    def test_line_spacing(self):
        #ExStart
        #ExFor:Font.line_spacing
        #ExSummary:Shows how to get a font's line spacing, in points.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Set different fonts for the DocumentBuilder and verify their line spacing.
        builder.font.name = 'Calibri'
        self.assertEqual(14.6484375, builder.font.line_spacing)
        builder.font.name = 'Times New Roman'
        self.assertEqual(13.798828125, builder.font.line_spacing)
        #ExEnd

    def test_has_dml_effect(self):
        #ExStart
        #ExFor:Font.has_dml_effect(TextDmlEffect)
        #ExFor:TextDmlEffect
        #ExSummary:Shows how to check if a run displays a DrawingML text effect.
        doc = aw.Document(file_name=MY_DIR + 'DrawingML text effects.docx')
        runs = doc.first_section.body.first_paragraph.runs
        self.assertTrue(runs[0].font.has_dml_effect(aw.TextDmlEffect.SHADOW))
        self.assertTrue(runs[1].font.has_dml_effect(aw.TextDmlEffect.SHADOW))
        self.assertTrue(runs[2].font.has_dml_effect(aw.TextDmlEffect.REFLECTION))
        self.assertTrue(runs[3].font.has_dml_effect(aw.TextDmlEffect.EFFECT_3D))
        self.assertTrue(runs[4].font.has_dml_effect(aw.TextDmlEffect.FILL))
        #ExEnd

    def test_set_emphasis_mark(self):
        for emphasis_mark in [aw.EmphasisMark.NONE, aw.EmphasisMark.OVER_COMMA, aw.EmphasisMark.OVER_SOLID_CIRCLE, aw.EmphasisMark.OVER_WHITE_CIRCLE, aw.EmphasisMark.UNDER_SOLID_CIRCLE]:
            #ExStart
            #ExFor:EmphasisMark
            #ExFor:Font.emphasis_mark
            #ExSummary:Shows how to add additional character rendered above/below the glyph-character.
            builder = aw.DocumentBuilder()
            # Possible types of emphasis mark:
            # https:#apireference.aspose.com/words/net/aspose.words/emphasismark
            builder.font.emphasis_mark = emphasis_mark
            builder.write('Emphasis text')
            builder.writeln()
            builder.font.clear_formatting()
            builder.write('Simple text')
            builder.document.save(file_name=ARTIFACTS_DIR + 'Fonts.SetEmphasisMark.docx')
        #ExEnd

    def test_theme_fonts_colors(self):
        #ExStart
        #ExFor:Font.theme_font
        #ExFor:Font.theme_font_ascii
        #ExFor:Font.theme_font_bi
        #ExFor:Font.theme_font_far_east
        #ExFor:Font.theme_font_other
        #ExFor:Font.theme_color
        #ExFor:ThemeFont
        #ExFor:ThemeColor
        #ExSummary:Shows how to work with theme fonts and colors.
        doc = aw.Document()
        # Define fonts for languages uses by default.
        doc.theme.minor_fonts.latin = 'Algerian'
        doc.theme.minor_fonts.east_asian = 'Aharoni'
        doc.theme.minor_fonts.complex_script = 'Andalus'
        font = doc.styles.get_by_name('Normal').font
        print('Originally the Normal style theme color is: {0} and RGB color is: {1}\n'.format(font.theme_color, font.color))
        # We can use theme font and color instead of default values.
        font.theme_font = aw.themes.ThemeFont.MINOR
        font.theme_color = aw.themes.ThemeColor.ACCENT2
        self.assertEqual(aw.themes.ThemeFont.MINOR, font.theme_font)
        self.assertEqual('Algerian', font.name)
        self.assertEqual(aw.themes.ThemeFont.MINOR, font.theme_font_ascii)
        self.assertEqual('Algerian', font.name_ascii)
        self.assertEqual(aw.themes.ThemeFont.MINOR, font.theme_font_bi)
        self.assertEqual('Andalus', font.name_bi)
        self.assertEqual(aw.themes.ThemeFont.MINOR, font.theme_font_far_east)
        self.assertEqual('Aharoni', font.name_far_east)
        self.assertEqual(aw.themes.ThemeFont.MINOR, font.theme_font_other)
        self.assertEqual('Algerian', font.name_other)
        self.assertEqual(aw.themes.ThemeColor.ACCENT2, font.theme_color)
        self.assertEqual(aspose.pydrawing.Color.empty(), font.color)
        # There are several ways of reset them font and color.
        # 1 -  By setting ThemeFont.None/ThemeColor.None:
        font.theme_font = aw.themes.ThemeFont.NONE
        font.theme_color = aw.themes.ThemeColor.NONE
        self.assertEqual(aw.themes.ThemeFont.NONE, font.theme_font)
        self.assertEqual('Algerian', font.name)
        self.assertEqual(aw.themes.ThemeFont.NONE, font.theme_font_ascii)
        self.assertEqual('Algerian', font.name_ascii)
        self.assertEqual(aw.themes.ThemeFont.NONE, font.theme_font_bi)
        self.assertEqual('Andalus', font.name_bi)
        self.assertEqual(aw.themes.ThemeFont.NONE, font.theme_font_far_east)
        self.assertEqual('Aharoni', font.name_far_east)
        self.assertEqual(aw.themes.ThemeFont.NONE, font.theme_font_other)
        self.assertEqual('Algerian', font.name_other)
        self.assertEqual(aw.themes.ThemeColor.NONE, font.theme_color)
        self.assertEqual(aspose.pydrawing.Color.empty(), font.color)
        # 2 -  By setting non-theme font/color names:
        font.name = 'Arial'
        font.color = aspose.pydrawing.Color.blue
        self.assertEqual(aw.themes.ThemeFont.NONE, font.theme_font)
        self.assertEqual('Arial', font.name)
        self.assertEqual(aw.themes.ThemeFont.NONE, font.theme_font_ascii)
        self.assertEqual('Arial', font.name_ascii)
        self.assertEqual(aw.themes.ThemeFont.NONE, font.theme_font_bi)
        self.assertEqual('Arial', font.name_bi)
        self.assertEqual(aw.themes.ThemeFont.NONE, font.theme_font_far_east)
        self.assertEqual('Arial', font.name_far_east)
        self.assertEqual(aw.themes.ThemeFont.NONE, font.theme_font_other)
        self.assertEqual('Arial', font.name_other)
        self.assertEqual(aw.themes.ThemeColor.NONE, font.theme_color)
        self.assertEqual(aspose.pydrawing.Color.blue.to_argb(), font.color.to_argb())
        #ExEnd

    def test_create_themed_style(self):
        #ExStart
        #ExFor:Font.theme_font
        #ExFor:Font.theme_color
        #ExFor:Font.tint_and_shade
        #ExFor:ThemeFont
        #ExFor:ThemeColor
        #ExSummary:Shows how to create and use themed style.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln()
        # Create some style with theme font properties.
        style = doc.styles.add(aw.StyleType.PARAGRAPH, 'ThemedStyle')
        style.font.theme_font = aw.themes.ThemeFont.MAJOR
        style.font.theme_color = aw.themes.ThemeColor.ACCENT5
        style.font.tint_and_shade = 0.3
        builder.paragraph_format.style_name = 'ThemedStyle'
        builder.writeln('Text with themed style')
        #ExEnd
        run = builder.current_paragraph.previous_sibling.as_paragraph().first_child.as_run()
        self.assertEqual(aw.themes.ThemeFont.MAJOR, run.font.theme_font)
        self.assertEqual('Times New Roman', run.font.name)
        self.assertEqual(aw.themes.ThemeFont.MAJOR, run.font.theme_font_ascii)
        self.assertEqual('Times New Roman', run.font.name_ascii)
        self.assertEqual(aw.themes.ThemeFont.MAJOR, run.font.theme_font_bi)
        self.assertEqual('Times New Roman', run.font.name_bi)
        self.assertEqual(aw.themes.ThemeFont.MAJOR, run.font.theme_font_far_east)
        self.assertEqual('Times New Roman', run.font.name_far_east)
        self.assertEqual(aw.themes.ThemeFont.MAJOR, run.font.theme_font_other)
        self.assertEqual('Times New Roman', run.font.name_other)
        self.assertEqual(aw.themes.ThemeColor.ACCENT5, run.font.theme_color)
        self.assertEqual(aspose.pydrawing.Color.empty(), run.font.color)

    def test_font_info_embedding_licensing_rights(self):
        #ExStart:FontInfoEmbeddingLicensingRights
        #ExFor:FontInfo.embedding_licensing_rights
        #ExFor:FontEmbeddingUsagePermissions
        #ExFor:FontEmbeddingLicensingRights
        #ExFor:FontEmbeddingLicensingRights.embedding_usage_permissions
        #ExFor:FontEmbeddingLicensingRights.bitmap_embedding_only
        #ExFor:FontEmbeddingLicensingRights.no_subsetting
        #ExSummary:Shows how to get license rights information for embedded fonts (FontInfo).
        doc = aw.Document(file_name=MY_DIR + 'Embedded font rights.docx')
        # Get the list of document fonts.
        font_infos = doc.font_infos
        for font_info in font_infos:
            if font_info.embedding_licensing_rights != None:
                print(font_info.embedding_licensing_rights.embedding_usage_permissions)
                print(font_info.embedding_licensing_rights.bitmap_embedding_only)
                print(font_info.embedding_licensing_rights.no_subsetting)
        #ExEnd:FontInfoEmbeddingLicensingRights

    def test_physical_font_info_embedding_licensing_rights(self):
        #ExStart:PhysicalFontInfoEmbeddingLicensingRights
        #ExFor:PhysicalFontInfo.embedding_licensing_rights
        #ExSummary:Shows how to get license rights information for embedded fonts (PhysicalFontInfo).
        settings = aw.fonts.FontSettings.default_instance
        source = settings.get_fonts_sources()[0]
        # Get the list of available fonts.
        font_infos = source.get_available_fonts()
        for font_info in font_infos:
            if font_info.embedding_licensing_rights != None:
                print(font_info.embedding_licensing_rights.embedding_usage_permissions)
                print(font_info.embedding_licensing_rights.bitmap_embedding_only)
                print(font_info.embedding_licensing_rights.no_subsetting)
        #ExEnd:PhysicalFontInfoEmbeddingLicensingRights

    def test_number_spacing(self):
        #ExStart:NumberSpacing
        #ExFor:Font.number_spacing
        #ExFor:NumSpacing
        #ExSummary:Shows how to set the spacing type of the numeral.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # This effect is only supported in newer versions of MS Word.
        doc.compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2019)
        builder.write('1 ')
        builder.write('This is an example')
        run = doc.first_section.body.first_paragraph.runs[0]
        if run.font.number_spacing == aw.NumSpacing.DEFAULT:
            run.font.number_spacing = aw.NumSpacing.PROPORTIONAL
        doc.save(file_name=ARTIFACTS_DIR + 'Fonts.NumberSpacing.docx')
        #ExEnd:NumberSpacing
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Fonts.NumberSpacing.docx')
        run = doc.first_section.body.first_paragraph.runs[0]
        self.assertEqual(aw.NumSpacing.PROPORTIONAL, run.font.number_spacing)

    def test_get_document_fonts(self):
        #ExStart
        #ExFor:FontInfoCollection
        #ExFor:DocumentBase.font_infos
        #ExFor:FontInfo
        #ExFor:FontInfo.name
        #ExFor:FontInfo.is_true_type
        #ExSummary:Shows how to print the details of what fonts are present in a document.
        doc = aw.Document(MY_DIR + 'Embedded font.docx')
        all_fonts = doc.font_infos
        self.assertEqual(5, all_fonts.count)  #ExSkip
        # Print all the used and unused fonts in the document.
        for i in range(all_fonts.count):
            print(f'Font index #{i}')
            print(f'\tName: {all_fonts[i].name}')
            print(f"\tIs {('' if all_fonts[i].is_true_type else 'not ')}a TrueType font")
        #ExEnd

    def test_locale_id(self):
        #ExStart
        #ExFor:Font.locale_id
        #ExSummary:Shows how to set the locale of the text that we are adding with a document builder.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        # If we set the font's locale to English and insert some Russian text,
        # the English locale spell checker will not recognize the text and detect it as a spelling error.
        builder.font.locale_id = 1033  # en-US
        builder.writeln('Привет!')
        # Set a matching locale for the text that we are about to add to apply the appropriate spell checker.
        builder.font.locale_id = 1049  # ru-RU
        builder.writeln('Привет!')
        doc.save(ARTIFACTS_DIR + 'Font.locale_id.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Font.locale_id.docx')
        run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('Привет!', run.get_text().strip())
        self.assertEqual(1033, run.font.locale_id)
        run = doc.first_section.body.paragraphs[1].runs[0]
        self.assertEqual('Привет!', run.get_text().strip())
        self.assertEqual(1049, run.font.locale_id)

    def test_bidi(self):
        #ExStart
        #ExFor:Font.bidi
        #ExFor:Font.name_bi
        #ExFor:Font.size_bi
        #ExFor:Font.italic_bi
        #ExFor:Font.bold_bi
        #ExFor:Font.locale_id_bi
        #ExSummary:Shows how to define separate sets of font settings for right-to-left, and right-to-left text.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        # Define a set of font settings for left-to-right text.
        builder.font.name = 'Courier New'
        builder.font.size = 16
        builder.font.italic = False
        builder.font.bold = False
        builder.font.locale_id = 1033  # en-US
        # Define another set of font settings for right-to-left text.
        builder.font.name_bi = 'Andalus'
        builder.font.size_bi = 24
        builder.font.italic_bi = True
        builder.font.bold_bi = True
        builder.font.locale_id_bi = 4096  # ar-AR
        # We can use the "bidi" flag to indicate whether the text we are about to add
        # with the document builder is right-to-left. When we add text with this flag set to True,
        # it will be formatted using the right-to-left set of font settings.
        builder.font.bidi = True
        builder.write('مرحبًا')
        # Set the flag to "False", and then add left-to-right text.
        # The document builder will format these using the left-to-right set of font settings.
        builder.font.bidi = False
        builder.write(' Hello world!')
        doc.save(ARTIFACTS_DIR + 'Font.bidi.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Font.bidi.docx')
        for run in doc.first_section.body.paragraphs[0].runs:
            run = run.as_run()
            idx = doc.first_section.body.paragraphs[0].index_of(run)
            if idx == 0:
                self.assertEqual('مرحبًا', run.get_text().strip())
                self.assertTrue(run.font.bidi)
            elif idx == 1:
                self.assertEqual('Hello world!', run.get_text().strip())
                self.assertFalse(run.font.bidi)
            self.assertEqual(1033, run.font.locale_id)
            self.assertEqual(16, run.font.size)
            self.assertFalse(run.font.italic)
            self.assertFalse(run.font.bold)
            self.assertEqual(1025, run.font.locale_id_bi)
            self.assertEqual(24, run.font.size_bi)
            self.assertEqual('Andalus', run.font.name_bi)
            self.assertTrue(run.font.italic_bi)
            self.assertTrue(run.font.bold_bi)

    def test_far_east(self):
        #ExStart
        #ExFor:Font.name_far_east
        #ExFor:Font.locale_id_far_east
        #ExSummary:Shows how to insert and format text in a Far East language.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        # Specify font settings that the document builder will apply to any text that it inserts.
        builder.font.name = 'Courier New'
        builder.font.locale_id = 1033  # en-US
        # Name "FarEast" equivalents for our font and locale.
        # If the builder inserts Asian characters with this Font configuration, then each run that contains
        # these characters will display them using the "FarEast" font/locale instead of the default.
        # This could be useful when a western font does not have ideal representations for Asian characters.
        builder.font.name_far_east = 'SimSun'
        builder.font.locale_id_far_east = 2052  # zh-CN
        # This text will be displayed in the default font/locale.
        builder.writeln('Hello world!')
        # Since these are Asian characters, this run will apply our "FarEast" font/locale equivalents.
        builder.writeln('你好世界')
        doc.save(ARTIFACTS_DIR + 'Font.far_east.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'Font.far_east.docx')
        run = doc.first_section.body.paragraphs[0].runs[0]
        self.assertEqual('Hello world!', run.get_text().strip())
        self.assertEqual(1033, run.font.locale_id)
        self.assertEqual('Courier New', run.font.name)
        self.assertEqual(2052, run.font.locale_id_far_east)
        self.assertEqual('SimSun', run.font.name_far_east)
        run = doc.first_section.body.paragraphs[1].runs[0]
        self.assertEqual('你好世界', run.get_text().strip())
        self.assertEqual(1033, run.font.locale_id)
        self.assertEqual('SimSun', run.font.name)
        self.assertEqual(2052, run.font.locale_id_far_east)
        self.assertEqual('SimSun', run.font.name_far_east)

    def test_get_available_fonts(self):
        #ExStart
        #ExFor:PhysicalFontInfo
        #ExFor:FontSourceBase.get_available_fonts
        #ExFor:PhysicalFontInfo.font_family_name
        #ExFor:PhysicalFontInfo.full_font_name
        #ExFor:PhysicalFontInfo.version
        #ExFor:PhysicalFontInfo.file_path
        #ExSummary:Shows how to list available fonts.
        # Configure Aspose.Words to source fonts from a custom folder, and then print every available font.
        folder_font_source = [aw.fonts.FolderFontSource(FONTS_DIR, True)]
        for font_info in folder_font_source[0].get_available_fonts():
            print('FontFamilyName :', font_info.font_family_name)
            print('FullFontName   :', font_info.full_font_name)
            print('Version  :', font_info.version)
            print('FilePath :', font_info.file_path)
            print()
        #ExEnd
        self.assertEqual(len(folder_font_source[0].get_available_fonts()), len(glob.glob(FONTS_DIR + '**/*.ttf', recursive=True) + glob.glob(FONTS_DIR + '**/*.otf', recursive=True)))

    def test_get_font_info_from_file(self):
        #ExStart
        #ExFor:FontFamily
        #ExFor:FontPitch
        #ExFor:FontInfo.alt_name
        #ExFor:FontInfo.charset
        #ExFor:FontInfo.family
        #ExFor:FontInfo.panose
        #ExFor:FontInfo.pitch
        #ExFor:FontInfoCollection.__iter__
        #ExSummary:Shows how to access and print details of each font in a document.
        doc = aw.Document(MY_DIR + 'Document.docx')
        for font_info in doc.font_infos:
            if font_info is not None:
                print('Font name: ' + font_info.name)
                # Alt names are usually blank.
                print('Alt name:', font_info.alt_name)
                print('\t- Family:', font_info.family)
                print('\t-', 'Is TrueType' if font_info.is_true_type else 'Is not TrueType')
                print('\t- Pitch:', font_info.pitch)
                print('\t- Charset:', font_info.charset)
                print('\t- Panose:')
                print('\t\tFamily Kind:', font_info.panose[0])
                print('\t\tSerif Style:', font_info.panose[1])
                print('\t\tWeight:', font_info.panose[2])
                print('\t\tProportion:', font_info.panose[3])
                print('\t\tContrast:', font_info.panose[4])
                print('\t\tStroke Variation:', font_info.panose[5])
                print('\t\tArm Style:', font_info.panose[6])
                print('\t\tLetterform:', font_info.panose[7])
                print('\t\tMidline:', font_info.panose[8])
                print('\t\tX-Height:', font_info.panose[9])
        #ExEnd
        self.assertEqual(bytes([2, 15, 5, 2, 2, 2, 4, 3, 2, 4]), doc.font_infos.get_by_name('Calibri').panose)
        self.assertEqual(bytes([2, 15, 3, 2, 2, 2, 4, 3, 2, 4]), doc.font_infos.get_by_name('Calibri Light').panose)
        self.assertEqual(bytes([2, 2, 6, 3, 5, 4, 5, 2, 3, 4]), doc.font_infos.get_by_name('Times New Roman').panose)

    @unittest.skipUnless(sys.platform.startswith('win'), 'requires Windows')
    def test_check_scan_user_fonts_folder(self):
        user_profile = pathlib.Path(os.environ['USERPROFILE'])
        current_user_fonts_folder = user_profile.joinpath('AppData\\Local\\Microsoft\\Windows\\Fonts')
        if len(list(current_user_fonts_folder.glob('*.ttf'))) > 0:
            # On Windows 10 fonts may be installed either into system folder "%windir%\fonts" for all users
            # or into user folder "%userprofile%\AppData\Local\Microsoft\Windows\Fonts" for current user.
            system_font_source = awfonts.SystemFontSource()
            current_user_fonts = filter(lambda font: font.find('\\AppData\\Local\\Microsoft\\Windows\\Fonts') != -1, system_font_source.get_system_font_folders())
            self.assertTrue(len(list(current_user_fonts)) > 0, 'Fonts did not install to the user font folder')

    def _test_remove_hidden_content(self, doc: aw.Document):
        self.assertEqual(20, doc.get_child_nodes(aw.NodeType.PARAGRAPH, True).count)
        self.assertEqual(1, doc.get_child_nodes(aw.NodeType.TABLE, True).count)
        for node in doc.get_child_nodes(aw.NodeType.ANY, True):
            if node in aw.fields.FieldStart:
                self.assertFalse(node.as_field_start().font.hidden)
            elif node is aw.fields.FieldEnd:
                self.assertFalse(node.as_field_end().font.hidden)
            elif node is aw.fields.FieldSeparator:
                self.assertFalse(node.as_field_separator().font.hidden)
            elif node is aw.Run:
                self.assertFalse(node.as_run().font.hidden)
            elif node is aw.Paragraph:
                self.assertFalse(node.as_paragraph().paragraph_break_font.hidden)
            elif node is aw.fields.FormField:
                self.assertFalse(node.as_form_field().font.hidden)
            elif node is aw.aspose.pydrawing.GroupShape:
                self.assertFalse(node.as_group_shape().font.hidden)
            elif node is aw.aspose.pydrawing.Shape:
                self.assertFalse(node.as_shape().font.hidden)
            elif node is aw.Comment:
                self.assertFalse(node.as_comment().font.hidden)
            elif node is aw.Footnote:
                self.assertFalse(node.as_footnote().font.hidden)
            elif node is aw.SpecialChar:
                self.assertFalse(node.as_special_char().font.hidden)

    def _test_remove_hidden_content(self, doc: aw.Document):
        self.assertEqual(20, doc.get_child_nodes(aw.NodeType.PARAGRAPH, True).count)
        self.assertEqual(1, doc.get_child_nodes(aw.NodeType.TABLE, True).count)
        for node in doc.get_child_nodes(aw.NodeType.ANY, True):
            if node in aw.fields.FieldStart:
                self.assertFalse(node.as_field_start().font.hidden)
            elif node is aw.fields.FieldEnd:
                self.assertFalse(node.as_field_end().font.hidden)
            elif node is aw.fields.FieldSeparator:
                self.assertFalse(node.as_field_separator().font.hidden)
            elif node is aw.Run:
                self.assertFalse(node.as_run().font.hidden)
            elif node is aw.Paragraph:
                self.assertFalse(node.as_paragraph().paragraph_break_font.hidden)
            elif node is aw.fields.FormField:
                self.assertFalse(node.as_form_field().font.hidden)
            elif node is aw.aspose.pydrawing.GroupShape:
                self.assertFalse(node.as_group_shape().font.hidden)
            elif node is aw.aspose.pydrawing.Shape:
                self.assertFalse(node.as_shape().font.hidden)
            elif node is aw.Comment:
                self.assertFalse(node.as_comment().font.hidden)
            elif node is aw.Footnote:
                self.assertFalse(node.as_footnote().font.hidden)
            elif node is aw.SpecialChar:
                self.assertFalse(node.as_special_char().font.hidden)