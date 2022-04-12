import aspose.words as aw
import aspose.pydrawing as drawing
from docs_examples_base import DocsExamplesBase, MY_DIR, ARTIFACTS_DIR

class WorkingWithFonts(DocsExamplesBase):

    def test_font_formatting(self):

        #ExStart:WriteAndFont
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        font = builder.font
        font.size = 16
        font.bold = True
        font.color = drawing.Color.blue
        font.name = "Arial"
        font.underline = aw.Underline.DASH

        builder.write("Sample text.")

        doc.save(ARTIFACTS_DIR + "WorkingWithFonts.font_formatting.docx")
        #ExEnd:WriteAndFont

    def test_get_font_line_spacing(self):

        #ExStart:GetFontLineSpacing
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.font.name = "Calibri"
        builder.writeln("qText")

        font = builder.document.first_section.body.first_paragraph.runs[0].font
        print(f"lineSpacing = {font.line_spacing}")
        #ExEnd:GetFontLineSpacing

    def test_check_dml_text_effect(self):

        #ExStart:CheckDMLTextEffect
        doc = aw.Document(MY_DIR + "DrawingML text effects.docx")

        runs = doc.first_section.body.first_paragraph.runs
        run_font = runs[0].font

        # One run might have several Dml text effects applied.
        print(run_font.has_dml_effect(aw.TextDmlEffect.SHADOW))
        print(run_font.has_dml_effect(aw.TextDmlEffect.EFFECT_3D))
        print(run_font.has_dml_effect(aw.TextDmlEffect.REFLECTION))
        print(run_font.has_dml_effect(aw.TextDmlEffect.OUTLINE))
        print(run_font.has_dml_effect(aw.TextDmlEffect.FILL))
        #ExEnd:CheckDMLTextEffect

    def test_set_font_formatting(self):

        #ExStart:DocumentBuilderSetFontFormatting
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        font = builder.font
        font.bold = True
        font.color = drawing.Color.dark_blue
        font.italic = True
        font.name = "Arial"
        font.size = 24
        font.spacing = 5
        font.underline = aw.Underline.DOUBLE

        builder.writeln("I'm a very nice formatted string.")

        doc.save(ARTIFACTS_DIR + "WorkingWithFonts.set_font_formatting.docx")
        #ExEnd:DocumentBuilderSetFontFormatting

    def test_set_font_emphasis_mark(self):

        #ExStart:SetFontEmphasisMark
        document = aw.Document()
        builder = aw.DocumentBuilder(document)

        builder.font.emphasis_mark = aw.EmphasisMark.UNDER_SOLID_CIRCLE

        builder.write("Emphasis text")
        builder.writeln()
        builder.font.clear_formatting()
        builder.write("Simple text")

        document.save(ARTIFACTS_DIR + "WorkingWithFonts.set_font_emphasis_mark.docx")
        #ExEnd:SetFontEmphasisMark

    def test_set_fonts_folders(self):

        #ExStart:SetFontsFolders
        aw.fonts.FontSettings.default_instance.set_fonts_sources([
            aw.fonts.SystemFontSource(),
            aw.fonts.FolderFontSource("C:\\MyFonts\\", True)])

        doc = aw.Document(MY_DIR + "Rendering.docx")
        doc.save(ARTIFACTS_DIR + "WorkingWithFonts.set_fonts_folders.pdf")
        #ExEnd:SetFontsFolders

    def test_enable_disable_font_substitution(self):

        #ExStart:EnableDisableFontSubstitution
        doc = aw.Document(MY_DIR + "Rendering.docx")

        font_settings = aw.fonts.FontSettings()
        font_settings.substitution_settings.default_font_substitution.default_font_name = "Arial"
        font_settings.substitution_settings.font_info_substitution.enabled = False

        doc.font_settings = font_settings

        doc.save(ARTIFACTS_DIR + "WorkingWithFonts.enable_disable_font_substitution.pdf")
        #ExEnd:EnableDisableFontSubstitution

    def test_set_font_fallback_settings(self):

        #ExStart:SetFontFallbackSettings
        doc = aw.Document(MY_DIR + "Rendering.docx")

        font_settings = aw.fonts.FontSettings()
        font_settings.fallback_settings.load(MY_DIR + "Font fallback rules.xml")

        doc.font_settings = font_settings

        doc.save(ARTIFACTS_DIR + "WorkingWithFonts.set_font_fallback_settings.pdf")
        #ExEnd:SetFontFallbackSettings

    def test_noto_fallback_settings(self):

        #ExStart:SetPredefinedFontFallbackSettings
        doc = aw.Document(MY_DIR + "Rendering.docx")

        font_settings = aw.fonts.FontSettings()
        font_settings.fallback_settings.load_noto_fallback_settings()

        doc.font_settings = font_settings

        doc.save(ARTIFACTS_DIR + "WorkingWithFonts.noto_fallback_settings.pdf")
        #ExEnd:SetPredefinedFontFallbackSettings

    def test_set_fonts_folders_default_instance(self):

        #ExStart:SetFontsFoldersDefaultInstance
        aw.fonts.FontSettings.default_instance.set_fonts_folder("C:\\MyFonts\\", True)
        #ExEnd:SetFontsFoldersDefaultInstance

        doc = aw.Document(MY_DIR + "Rendering.docx")
        doc.save(ARTIFACTS_DIR + "WorkingWithFonts.set_fonts_folders_default_instance.pdf")

    def test_set_fonts_folders_multiple_folders(self):

        #ExStart:SetFontsFoldersMultipleFolders
        doc = aw.Document(MY_DIR + "Rendering.docx")

        font_settings = aw.fonts.FontSettings()
        # Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
        # fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.get_font_sources and
        # FontSettings.set_font_sources instead.
        font_settings.set_fonts_folders(["C:\\MyFonts\\", "D:\\Misc\\Fonts\\"], True)

        doc.font_settings = font_settings

        doc.save(ARTIFACTS_DIR + "WorkingWithFonts.set_fonts_folders_multiple_folders.pdf")
        #ExEnd:SetFontsFoldersMultipleFolders

    def test_set_fonts_folders_system_and_custom_folder(self):

        #ExStart:SetFontsFoldersSystemAndCustomFolder
        doc = aw.Document(MY_DIR + "Rendering.docx")

        font_settings = aw.fonts.FontSettings()
        # Retrieve the array of environment-dependent font sources that are searched by default.
        # For example this will contain a "Windows\Fonts\" source on a Windows machines.
        # We add this array to a new List to make adding or removing font entries much easier.
        font_sources = font_settings.get_fonts_sources()

        # Add a new folder source which will instruct Aspose.words to search the following folder for fonts.
        folder_font_source = aw.fonts.FolderFontSource("C:\\MyFonts\\", True)

        # Add the custom folder which contains our fonts to the list of existing font sources.
        updated_font_sources = list(font_sources)
        updated_font_sources.append(folder_font_source)

        font_settings.set_fonts_sources(updated_font_sources)

        doc.font_settings = font_settings

        doc.save(ARTIFACTS_DIR + "WorkingWithFonts.set_fonts_folders_system_and_custom_folder.pdf")
        #ExEnd:SetFontsFoldersSystemAndCustomFolder

    def test_set_fonts_folders_with_priority(self):

        #ExStart:SetFontsFoldersWithPriority
        aw.fonts.FontSettings.default_instance.set_fonts_sources([
            aw.fonts.SystemFontSource(),
            aw.fonts.FolderFontSource("C:\\MyFonts\\", True, 1)])
        #ExEnd:SetFontsFoldersWithPriority

        doc = aw.Document(MY_DIR + "Rendering.docx")
        doc.save(ARTIFACTS_DIR + "WorkingWithFonts.set_fonts_folders_with_priority.pdf")

    def test_set_true_type_fonts_folder(self):

        #ExStart:SetTrueTypeFontsFolder
        doc = aw.Document(MY_DIR + "Rendering.docx")

        font_settings = aw.fonts.FontSettings()
        # Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
        # Fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.get_font_sources and
        # FontSettings.set_font_sources instead
        font_settings.set_fonts_folder("C:\\MyFonts\\", False)
        # Set font settings
        doc.font_settings = font_settings

        doc.save(ARTIFACTS_DIR + "WorkingWithFonts.set_true_type_fonts_folder.pdf")
        #ExEnd:SetTrueTypeFontsFolder

    def test_specify_default_font_when_rendering(self):

        #ExStart:SpecifyDefaultFontWhenRendering
        doc = aw.Document(MY_DIR + "Rendering.docx")

        font_settings = aw.fonts.FontSettings()
        # If the default font defined here cannot be found during rendering then
        # the closest font on the machine is used instead.
        font_settings.substitution_settings.default_font_substitution.default_font_name = "Arial Unicode MS"

        doc.font_settings = font_settings

        doc.save(ARTIFACTS_DIR + "WorkingWithFonts.specify_default_font_when_rendering.pdf")
        #ExEnd:SpecifyDefaultFontWhenRendering

    def test_font_settings_with_load_options(self):

        #ExStart:FontSettingsWithLoadOptions
        font_settings = aw.fonts.FontSettings()

        substitution_rule = font_settings.substitution_settings.table_substitution
        # If "UnknownFont1" font family is not available then substitute it by "Comic Sans MS"
        substitution_rule.add_substitutes("UnknownFont1", ["Comic Sans MS"])

        load_options = aw.loading.LoadOptions()
        load_options.font_settings = font_settings

        doc = aw.Document(MY_DIR + "Rendering.docx", load_options)
        #ExEnd:FontSettingsWithLoadOptions

    def test_set_fonts_folder(self):

        #ExStart:SetFontsFolder
        font_settings = aw.fonts.FontSettings()
        font_settings.set_fonts_folder(MY_DIR + "Fonts", False)

        load_options = aw.loading.LoadOptions()
        load_options.font_settings = font_settings

        doc = aw.Document(MY_DIR + "Rendering.docx", load_options)
        #ExEnd:SetFontsFolder

    def test_font_settings_with_load_option(self):

        #ExStart:FontSettingsWithLoadOption
        load_options = aw.loading.LoadOptions()
        load_options.font_settings = aw.fonts.FontSettings()

        doc = aw.Document(MY_DIR + "Rendering.docx", load_options)
        #ExEnd:FontSettingsWithLoadOption

    def test_font_settings_default_instance(self):

        #ExStart:FontSettingsFontSource
        #ExStart:FontSettingsDefaultInstance
        font_settings = aw.fonts.FontSettings.default_instance
        #ExEnd:FontSettingsDefaultInstance
        font_settings.set_fonts_sources([
            aw.fonts.SystemFontSource(),
            aw.fonts.FolderFontSource("C:\\MyFonts\\", True)])
        #ExEnd:FontSettingsFontSource

        load_options = aw.loading.LoadOptions()
        load_options.font_settings = font_settings
        doc = aw.Document(MY_DIR + "Rendering.docx", load_options)

    def test_get_list_of_available_fonts(self):

        #ExStart:GetListOfAvailableFonts
        font_settings = aw.fonts.FontSettings()
        font_sources = font_settings.get_fonts_sources()

        # Add a new folder source which will instruct Aspose.words to search the following folder for fonts.
        folder_font_source = aw.fonts.FolderFontSource(MY_DIR, True)
        # Add the custom folder which contains our fonts to the list of existing font sources.
        updated_font_sources = list(font_sources)
        updated_font_sources.append(folder_font_source)

        for font_info in updated_font_sources[0].get_available_fonts():
            print("FontFamilyName: " + font_info.font_family_name)
            print("FullFontName: " + font_info.full_font_name)
            print("Version: " + font_info.version)
            print("FilePath: " + font_info.file_path)

        #ExEnd:GetListOfAvailableFonts

#    @unittest.skip("Interface implementation is not supported yet.")
#    def test_receive_notifications_of_fonts(self):
#
#        #ExStart:ReceiveNotificationsOfFonts
#        doc = aw.Document(MY_DIR + "Rendering.docx")
#
#        fontSettings = aw.fonts.FontSettings()
#
#        # We can choose the default font to use in the case of any missing fonts.
#        fontSettings.substitution_settings.default_font_substitution.default_font_name = "Arial"
#        # For testing we will set Aspose.words to look for fonts only in a folder which doesn't exist. Since Aspose.words won't
#        # find any fonts in the specified directory, then during rendering the fonts in the document will be subsuited with the default
#        # font specified under FontSettings.default_font_name. We can pick up on this subsuition using our callback.
#        fontSettings.set_fonts_folder(string.empty, false)
#
#        # Create a new class implementing IWarningCallback which collect any warnings produced during document save.
#        HandleDocumentWarnings callback = new HandleDocumentWarnings()
#
#        doc.warning_callback = callback
#        doc.font_settings = fontSettings
#
#        doc.save(ARTIFACTS_DIR + "WorkingWithFonts.receive_notifications_of_fonts.pdf")
#        #ExEnd:ReceiveNotificationsOfFonts
#
#    @unittest.skip("Interface implementation is not supported yet.")
#    def test_receive_warning_notification(self):
#
#        #ExStart:ReceiveWarningNotification
#        doc = aw.Document(MY_DIR + "Rendering.docx")
#
#        # When you call UpdatePageLayout the document is rendered in memory. Any warnings that occured during rendering
#        # are stored until the document save and then sent to the appropriate WarningCallback.
#        doc.update_page_layout()
#
#        HandleDocumentWarnings callback = new HandleDocumentWarnings()
#        doc.warning_callback = callback
#
#        # Even though the document was rendered previously, any save warnings are notified to the user during document save.
#        doc.save(ARTIFACTS_DIR + "WorkingWithFonts.receive_warning_notification.pdf")
#        #ExEnd:ReceiveWarningNotification
#
#
#    #ExStart:HandleDocumentWarnings
#    public class HandleDocumentWarnings: IWarningCallback
#
#        # <summary>
#        # Our callback only needs to implement the "Warning" method. This method is called whenever there is a
#        # Potential issue during document procssing. The callback can be set to listen for warnings generated
#        # during document load and/or document save.
#        # </summary>
#        public void Warning(WarningInfo info)
#
#            # We are only interested in fonts being substituted.
#            if (info.warning_type == WarningType.font_substitution)
#
#                print("Font substitution: " + info.description)
#
#
#
#    #ExEnd:HandleDocumentWarnings
#
#    #ExStart:ResourceSteamFontSourceExample
#    def test_resource_steam_font_source_example(self):
#
#        doc = aw.Document(MY_DIR + "Rendering.docx")
#
#        FontSettings.default_instance.set_fonts_sources(new FontSourceBase[]
#                new SystemFontSource(), new ResourceSteamFontSource() )
#
#        doc.save(ARTIFACTS_DIR + "WorkingWithFonts.set_fonts_folders.pdf")
#
#
#    internal class ResourceSteamFontSource: StreamFontSource
#
#        public override Stream OpenFontDataStream()
#
#            return Assembly.get_executing_assembly().get_manifest_resource_stream("resourceName")
#
#
#    #ExEnd:ResourceSteamFontSourceExample
#
#    #ExStart:GetSubstitutionWithoutSuffixes
#    def test_get_substitution_without_suffixes(self):
#
#        doc = aw.Document(MY_DIR + "Get substitution without suffixes.docx")
#
#        DocumentSubstitutionWarnings substitutionWarningHandler = new DocumentSubstitutionWarnings()
#        doc.warning_callback = substitutionWarningHandler
#
#        List<FontSourceBase> fontSources = new List<FontSourceBase>(FontSettings.default_instance.get_fonts_sources())
#
#        FolderFontSource folderFontSource = new FolderFontSource(docs_base.fonts_dir, true)
#        fontSources.add(folderFontSource)
#
#        FontSourceBase[] updatedFontSources = fontSources.to_array()
#        FontSettings.default_instance.set_fonts_sources(updatedFontSources)
#
#        doc.save(ARTIFACTS_DIR + "WorkingWithFonts.get_substitution_without_suffixes.pdf")
#
#        self.assertEqual(
#            "Font 'DINOT-Regular' has not been found. Using 'DINOT' font instead. Reason: font name substitution.",
#            substitutionWarningHandler.font_warnings[0].description)
#
#
#    public class DocumentSubstitutionWarnings: IWarningCallback
#
#        # <summary>
#        # Our callback only needs to implement the "Warning" method.
#        # This method is called whenever there is a potential issue during document processing.
#        # The callback can be set to listen for warnings generated during document load and/or document save.
#        # </summary>
#        public void Warning(WarningInfo info)
#
#            # We are only interested in fonts being substituted.
#            if (info.warning_type == WarningType.font_substitution)
#                FontWarnings.warning(info)
#
#
#        public WarningInfoCollection FontWarnings = new WarningInfoCollection()
#
#    #ExEnd:GetSubstitutionWithoutSuffixes
