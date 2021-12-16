import unittest
import io

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

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
        doc = aw.Document(MY_DIR + "Theme colors.docx")

        # The "theme" object gives us access to the document theme, a source of default fonts and colors.
        theme = doc.theme

        # Some styles, such as "Heading 1" and "Subtitle", will inherit these fonts.
        theme.major_fonts.latin = "Courier New"
        theme.minor_fonts.latin = "Agency FB"

        # Other languages may also have their custom fonts in this theme.
        self.assertEqual("", theme.major_fonts.complex_script)
        self.assertEqual("", theme.major_fonts.east_asian)
        self.assertEqual("", theme.minor_fonts.complex_script)
        self.assertEqual("", theme.minor_fonts.east_asian)

        # The "colors" property contains the color palette from Microsoft Word,
        # which appears when changing shading or font color.
        # Apply custom colors to the color palette so we have easy access to them in Microsoft Word
        # when we, for example, change the font color via "Home" -> "Font" -> "Font Color",
        # or insert a shape, and then set a color for it via "Shape Format" -> "Shape Styles".
        colors = theme.colors
        colors.dark1 = drawing.Color.midnight_blue
        colors.light1 = drawing.Color.pale_green
        colors.dark2 = drawing.Color.indigo
        colors.light2 = drawing.Color.khaki

        colors.accent1 = drawing.Color.orange_red
        colors.accent2 = drawing.Color.light_salmon
        colors.accent3 = drawing.Color.yellow
        colors.accent4 = drawing.Color.gold
        colors.accent5 = drawing.Color.blue_violet
        colors.accent6 = drawing.Color.dark_violet

        # Apply custom colors to hyperlinks in their clicked and un-clicked states.
        colors.hyperlink = drawing.Color.black
        colors.followed_hyperlink = drawing.Color.gray

        doc.save(ARTIFACTS_DIR + "Themes.CustomColorsAndFonts.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "Themes.CustomColorsAndFonts.docx")

        self.assertEqual(drawing.Color.orange_red.to_argb(), doc.theme.colors.accent1.to_argb())
        self.assertEqual(drawing.Color.midnight_blue.to_argb(), doc.theme.colors.dark1.to_argb())
        self.assertEqual(drawing.Color.gray.to_argb(), doc.theme.colors.followed_hyperlink.to_argb())
        self.assertEqual(drawing.Color.black.to_argb(), doc.theme.colors.hyperlink.to_argb())
        self.assertEqual(drawing.Color.pale_green.to_argb(), doc.theme.colors.light1.to_argb())

        self.assertEqual("", doc.theme.major_fonts.complex_script)
        self.assertEqual("", doc.theme.major_fonts.east_asian)
        self.assertEqual("Courier New", doc.theme.major_fonts.latin)

        self.assertEqual("", doc.theme.minor_fonts.complex_script)
        self.assertEqual("", doc.theme.minor_fonts.east_asian)
        self.assertEqual("Agency FB", doc.theme.minor_fonts.latin)
