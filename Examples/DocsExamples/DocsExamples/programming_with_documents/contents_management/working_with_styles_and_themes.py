import aspose.words as aw
import aspose.pydrawing as drawing
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class WorkingWithStylesAndThemes(DocsExamplesBase):

    def test_access_styles(self):

        #ExStart:AccessStyles
        doc = aw.Document()

        style_name = ""

        # Get styles collection from the document.
        styles = doc.styles
        for style in styles:
            if style_name == "":
                style_name = style.name
                print(style_name)
            else:
                style_name = style_name + ", " + style.name
                print(style_name)

        #ExEnd:AccessStyles

    def test_copy_styles(self):

        #ExStart:CopyStyles
        doc = aw.Document()
        target = aw.Document(MY_DIR + "Rendering.docx")

        target.copy_styles_from_template(doc)

        doc.save(ARTIFACTS_DIR + "WorkingWithStylesAndThemes.copy_styles.docx")
        #ExEnd:CopyStyles

    def test_get_theme_properties(self):

        #ExStart:GetThemeProperties
        doc = aw.Document()

        theme = doc.theme

        print(theme.major_fonts.latin)
        print(theme.minor_fonts.east_asian)
        print(theme.colors.accent1)
        #ExEnd:GetThemeProperties

    def test_set_theme_properties(self):

        #ExStart:SetThemeProperties
        doc = aw.Document()

        theme = doc.theme
        theme.minor_fonts.latin = "Times New Roman"
        theme.colors.hyperlink = drawing.Color.gold
        #ExEnd:SetThemeProperties

    def test_insert_style_separator(self):

        #ExStart:InsertStyleSeparator
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        para_style = builder.document.styles.add(aw.StyleType.PARAGRAPH, "MyParaStyle")
        para_style.font.bold = False
        para_style.font.size = 8
        para_style.font.name = "Arial"

        # Append text with "Heading 1" style.
        builder.paragraph_format.style_identifier = aw.StyleIdentifier.HEADING1
        builder.write("Heading 1")
        builder.insert_style_separator()

        # Append text with another style.
        builder.paragraph_format.style_name = para_style.name
        builder.write("This is text with some other formatting ")

        doc.save(ARTIFACTS_DIR + "WorkingWithStylesAndThemes.insert_style_separator.docx")
        #ExEnd:InsertStyleSeparator
