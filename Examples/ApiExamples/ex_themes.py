# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.pydrawing
import aspose.words as aw
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExThemes(ApiExampleBase):

    def test_custom_colors_and_fonts(self):
        #ExStart
        #ExFor:Document.theme
        #ExFor:Theme
        #ExFor:Theme.colors
        #ExFor:Theme.major_fonts
        #ExFor:Theme.minor_fonts
        #ExFor:ThemeColors
        #ExFor:ThemeColors.accent1
        #ExFor:ThemeColors.accent2
        #ExFor:ThemeColors.accent3
        #ExFor:ThemeColors.accent4
        #ExFor:ThemeColors.accent5
        #ExFor:ThemeColors.accent6
        #ExFor:ThemeColors.dark1
        #ExFor:ThemeColors.dark2
        #ExFor:ThemeColors.followed_hyperlink
        #ExFor:ThemeColors.hyperlink
        #ExFor:ThemeColors.light1
        #ExFor:ThemeColors.light2
        #ExFor:ThemeFonts
        #ExFor:ThemeFonts.complex_script
        #ExFor:ThemeFonts.east_asian
        #ExFor:ThemeFonts.latin
        #ExSummary:Shows how to set custom colors and fonts for themes.
        doc = aw.Document(file_name=MY_DIR + 'Theme colors.docx')
        # The "Theme" object gives us access to the document theme, a source of default fonts and colors.
        theme = doc.theme
        # Some styles, such as "Heading 1" and "Subtitle", will inherit these fonts.
        theme.major_fonts.latin = 'Courier New'
        theme.minor_fonts.latin = 'Agency FB'
        # Other languages may also have their custom fonts in this theme.
        self.assertEqual('', theme.major_fonts.complex_script)
        self.assertEqual('', theme.major_fonts.east_asian)
        self.assertEqual('', theme.minor_fonts.complex_script)
        self.assertEqual('', theme.minor_fonts.east_asian)
        # The "Colors" property contains the color palette from Microsoft Word,
        # which appears when changing shading or font color.
        # Apply custom colors to the color palette so we have easy access to them in Microsoft Word
        # when we, for example, change the font color via "Home" -> "Font" -> "Font Color",
        # or insert a shape, and then set a color for it via "Shape Format" -> "Shape Styles".
        colors = theme.colors
        colors.dark1 = aspose.pydrawing.Color.midnight_blue
        colors.light1 = aspose.pydrawing.Color.pale_green
        colors.dark2 = aspose.pydrawing.Color.indigo
        colors.light2 = aspose.pydrawing.Color.khaki
        colors.accent1 = aspose.pydrawing.Color.orange_red
        colors.accent2 = aspose.pydrawing.Color.light_salmon
        colors.accent3 = aspose.pydrawing.Color.yellow
        colors.accent4 = aspose.pydrawing.Color.gold
        colors.accent5 = aspose.pydrawing.Color.blue_violet
        colors.accent6 = aspose.pydrawing.Color.dark_violet
        # Apply custom colors to hyperlinks in their clicked and un-clicked states.
        colors.hyperlink = aspose.pydrawing.Color.black
        colors.followed_hyperlink = aspose.pydrawing.Color.gray
        doc.save(file_name=ARTIFACTS_DIR + 'Themes.CustomColorsAndFonts.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Themes.CustomColorsAndFonts.docx')
        self.assertEqual(aspose.pydrawing.Color.orange_red.to_argb(), doc.theme.colors.accent1.to_argb())
        self.assertEqual(aspose.pydrawing.Color.midnight_blue.to_argb(), doc.theme.colors.dark1.to_argb())
        self.assertEqual(aspose.pydrawing.Color.gray.to_argb(), doc.theme.colors.followed_hyperlink.to_argb())
        self.assertEqual(aspose.pydrawing.Color.black.to_argb(), doc.theme.colors.hyperlink.to_argb())
        self.assertEqual(aspose.pydrawing.Color.pale_green.to_argb(), doc.theme.colors.light1.to_argb())
        self.assertEqual('', doc.theme.major_fonts.complex_script)
        self.assertEqual('', doc.theme.major_fonts.east_asian)
        self.assertEqual('Courier New', doc.theme.major_fonts.latin)
        self.assertEqual('', doc.theme.minor_fonts.complex_script)
        self.assertEqual('', doc.theme.minor_fonts.east_asian)
        self.assertEqual('Agency FB', doc.theme.minor_fonts.latin)