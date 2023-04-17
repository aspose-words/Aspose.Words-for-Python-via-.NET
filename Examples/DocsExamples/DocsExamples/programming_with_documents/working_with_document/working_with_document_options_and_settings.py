import aspose.words as aw
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR
from aspose.pydrawing import Color

class WorkingWithDocumentOptionsAndSettings(DocsExamplesBase):

    def test_optimize_for_ms_word(self):

        #ExStart:OptimizeForMsWord
        doc = aw.Document(MY_DIR + "Document.docx")

        doc.compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2016)

        doc.save(ARTIFACTS_DIR + "WorkingWithDocumentOptionsAndSettings.optimize_for_ms_word.docx")
        #ExEnd:OptimizeForMsWord

    def test_show_grammatical_and_spelling_errors(self):

        #ExStart:ShowGrammaticalAndSpellingErrors
        doc = aw.Document(MY_DIR + "Document.docx")

        doc.show_grammatical_errors = True
        doc.show_spelling_errors = True

        doc.save(ARTIFACTS_DIR + "WorkingWithDocumentOptionsAndSettings.show_grammatical_and_spelling_errors.docx")
        #ExEnd:ShowGrammaticalAndSpellingErrors

    def test_cleanup_unused_styles_and_lists(self):

        #ExStart:CleanupUnusedStylesandLists
        doc = aw.Document(MY_DIR + "Unused styles.docx")

        # Combined with the built-in styles, the document now has eight styles.
        # A custom style is marked as "used" while there is any text within the document
        # formatted in that style. This means that the 4 styles we added are currently unused.
        print("Count of styles before Cleanup:", doc.styles.count)
        print("Count of lists before Cleanup:", doc.lists.count)

        # Cleans unused styles and lists from the document depending on given CleanupOptions.
        cleanup_options = aw.CleanupOptions()
        cleanup_options.unused_lists = False
        cleanup_options.unused_styles = True
        doc.cleanup(cleanup_options)

        print("Count of styles after Cleanup was decreased:", doc.styles.count)
        print("Count of lists after Cleanup is the same:", doc.lists.count)

        doc.save(ARTIFACTS_DIR + "WorkingWithDocumentOptionsAndSettings.cleanup_unused_styles_and_lists.docx")
        #ExEnd:CleanupUnusedStylesandLists

    def test_cleanup_duplicate_style(self):

        #ExStart:CleanupDuplicateStyle
        doc = aw.Document(MY_DIR + "Document.docx")

        # Count of styles before Cleanup.
        print(doc.styles.count)

        # Cleans duplicate styles from the document.
        options = aw.CleanupOptions()
        options.duplicate_style = True
        doc.cleanup(options)

        # Count of styles after Cleanup was decreased.
        print(doc.styles.count)

        doc.save(ARTIFACTS_DIR + "WorkingWithDocumentOptionsAndSettings.cleanup_duplicate_style.docx")
        #ExEnd:CleanupDuplicateStyle

    def test_view_options(self):

        #ExStart:SetViewOption
        doc = aw.Document(MY_DIR + "Document.docx")

        doc.view_options.view_type = aw.settings.ViewType.PAGE_LAYOUT
        doc.view_options.zoom_percent = 50

        doc.save(ARTIFACTS_DIR + "WorkingWithDocumentOptionsAndSettings.view_options.docx")
        #ExEnd:SetViewOption

    def test_document_page_setup(self):

        #ExStart:DocumentPageSetup
        doc = aw.Document(MY_DIR + "Document.docx")

        # Set the layout mode for a section allowing to define the document grid behavior.
        # Note that the Document Grid tab becomes visible in the Page Setup dialog of MS Word
        # if any Asian language is defined as editing language.
        doc.first_section.page_setup.layout_mode = aw.SectionLayoutMode.GRID
        doc.first_section.page_setup.characters_per_line = 30
        doc.first_section.page_setup.lines_per_page = 10

        doc.save(ARTIFACTS_DIR + "WorkingWithDocumentOptionsAndSettings.document_page_setup.docx")
        #ExEnd:DocumentPageSetup

    def test_add_japanese_as_editing_languages(self):

        #ExStart:AddJapaneseAsEditinglanguages
        load_options = aw.loading.LoadOptions()

        # Set language preferences that will be used when document is loading.
        load_options.language_preferences.add_editing_language(aw.loading.EditingLanguage.JAPANESE)
        #ExEnd:AddJapaneseAsEditinglanguages

        doc = aw.Document(MY_DIR + "No default editing language.docx", load_options)

        locale_id_far_east = doc.styles.default_font.locale_id_far_east
        if locale_id_far_east == aw.loading.EditingLanguage.JAPANESE:
            print("The document either has no any FarEast language set in defaults or it was set to Japanese originally.")
        else:
            print("The document default FarEast language was set to another than Japanese language originally, so it is not overridden.")

    def test_set_russian_as_default_editing_language(self):

        #ExStart:SetRussianAsDefaultEditingLanguage
        load_options = aw.loading.LoadOptions()
        load_options.language_preferences.default_editing_language = aw.loading.EditingLanguage.RUSSIAN

        doc = aw.Document(MY_DIR + "No default editing language.docx", load_options)

        locale_id = doc.styles.default_font.locale_id
        if locale_id == aw.loading.EditingLanguage.RUSSIAN:
            print("The document either has no any language set in defaults or it was set to Russian originally.")
        else:
            print("The document default language was set to another than Russian language originally, so it is not overridden.")
        #ExEnd:SetRussianAsDefaultEditingLanguage

    def test_set_page_setup_and_section_formatting(self):

        #ExStart:DocumentBuilderSetPageSetupAndSectionFormatting
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.page_setup.orientation = aw.Orientation.LANDSCAPE
        builder.page_setup.left_margin = 50
        builder.page_setup.paper_size = aw.PaperSize.PAPER_10X14

        doc.save(ARTIFACTS_DIR + "WorkingWithDocumentOptionsAndSettings.set_page_setup_and_section_formatting.docx")
        #ExEnd:DocumentBuilderSetPageSetupAndSectionFormatting

    def test_page_border_properties(self):
        #ExStart: PageBorderProperties
        doc = aw.Document()
        page_setup = doc.sections[0].page_setup

        page_setup.border_always_in_front = False
        page_setup.border_distance_from = aw.PageBorderDistanceFrom.PAGE_EDGE
        page_setup.border_applies_to = aw.PageBorderAppliesTo.FIRST_PAGE

        border = page_setup.borders.get_by_border_type(aw.BorderType.TOP)
        border.line_style = aw.LineStyle.SINGLE
        border.line_width = 30
        border.color = Color.blue
        border.distance_from_text = 0

        doc.save(ARTIFACTS_DIR + "WorkingWithDocumentOptionsAndSettings.PageBorderProperties.docx")
        #ExEnd: PageBorderProperties


    def test_line_grid_section_layoutMode(self):

        #ExStart: LineGridSectionLayoutMode
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Enable pitching, and then use it to set the number of lines per page in this section.
        # A large enough font size will push some lines down onto the next page to avoid overlapping characters.
        builder.page_setup.layout_mode = aw.SectionLayoutMode.LINE_GRID
        builder.page_setup.lines_per_page = 15
        builder.paragraph_format.snap_to_grid = True
        for i in range(30):
            builder.write(
                "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ")

        doc.save(ARTIFACTS_DIR + "WorkingWithDocumentOptionsAndSettings.LinesPerPage.docx")
        #ExEnd:LineGridSectionLayoutMode


