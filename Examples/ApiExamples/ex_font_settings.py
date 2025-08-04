# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import sys
import xml.etree.ElementTree as ET
import platform
import aspose.words as aw
import aspose.words.fonts
import aspose.words.loading
import io
import system_helper
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, FONTS_DIR, MY_DIR

class ExFontSettings(ApiExampleBase):

    def test_default_font_instance(self):
        #ExStart
        #ExFor:FontSettings.default_instance
        #ExSummary:Shows how to configure the default font settings instance.
        # Configure the default font settings instance to use the "Courier New" font
        # as a backup substitute when we attempt to use an unknown font.
        aw.fonts.FontSettings.default_instance.substitution_settings.default_font_substitution.default_font_name = 'Courier New'
        self.assertTrue(aw.fonts.FontSettings.default_instance.substitution_settings.default_font_substitution.enabled)
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.name = 'Non-existent font'
        builder.write('Hello world!')
        # This document does not have a FontSettings configuration. When we render the document,
        # the default FontSettings instance will resolve the missing font.
        # Aspose.Words will use "Courier New" to render text that uses the unknown font.
        self.assertIsNone(doc.font_settings)
        doc.save(file_name=ARTIFACTS_DIR + 'FontSettings.DefaultFontInstance.pdf')
        #ExEnd

    def test_default_font_name(self):
        #ExStart
        #ExFor:DefaultFontSubstitutionRule.default_font_name
        #ExSummary:Shows how to specify a default font.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.name = 'Arial'
        builder.writeln('Hello world!')
        builder.font.name = 'Arvo'
        builder.writeln('The quick brown fox jumps over the lazy dog.')
        font_sources = aw.fonts.FontSettings.default_instance.get_fonts_sources()
        # The font sources that the document uses contain the font "Arial", but not "Arvo".
        self.assertEqual(1, len(font_sources))
        self.assertTrue(any([f.full_font_name == 'Arial' for f in font_sources[0].get_available_fonts()]))
        self.assertFalse(any([f.full_font_name == 'Arvo' for f in font_sources[0].get_available_fonts()]))
        # Set the "DefaultFontName" property to "Courier New" to,
        # while rendering the document, apply that font in all cases when another font is not available.
        aw.fonts.FontSettings.default_instance.substitution_settings.default_font_substitution.default_font_name = 'Courier New'
        self.assertTrue(any([f.full_font_name == 'Courier New' for f in font_sources[0].get_available_fonts()]))
        # Aspose.Words will now use the default font in place of any missing fonts during any rendering calls.
        doc.save(file_name=ARTIFACTS_DIR + 'FontSettings.DefaultFontName.pdf')
        #ExEnd

    @unittest.skip('Discrepancy in assertion between Python and .Net')
    def test_enable_font_substitution(self):
        #ExStart
        #ExFor:FontInfoSubstitutionRule
        #ExFor:FontSubstitutionSettings.font_info_substitution
        #ExFor:LayoutOptions.keep_original_font_metrics
        #ExFor:IWarningCallback
        #ExFor:IWarningCallback.warning(WarningInfo)
        #ExFor:WarningInfo
        #ExFor:WarningInfo.description
        #ExFor:WarningInfo.warning_type
        #ExFor:WarningInfoCollection
        #ExFor:WarningInfoCollection.warning(WarningInfo)
        #ExFor:WarningInfoCollection.clear
        #ExFor:WarningType
        #ExFor:DocumentBase.warning_callback
        #ExSummary:Shows how to set the property for finding the closest match for a missing font from the available font sources.
        # Open a document that contains text formatted with a font that does not exist in any of our font sources.
        doc = aw.Document(file_name=MY_DIR + 'Missing font.docx')
        # Assign a callback for handling font substitution warnings.
        warning_collector = aw.WarningInfoCollection()
        doc.warning_callback = warning_collector
        # Set a default font name and enable font substitution.
        font_settings = aw.fonts.FontSettings()
        font_settings.substitution_settings.default_font_substitution.default_font_name = 'Arial'
        font_settings.substitution_settings.font_info_substitution.enabled = True
        # Original font metrics should be used after font substitution.
        doc.layout_options.keep_original_font_metrics = True
        # We will get a font substitution warning if we save a document with a missing font.
        doc.font_settings = font_settings
        doc.save(file_name=ARTIFACTS_DIR + 'FontSettings.EnableFontSubstitution.pdf')
        for info in warning_collector:
            if info.warning_type == aw.WarningType.FONT_SUBSTITUTION:
                print(info.description)
        #ExEnd
        # We can also verify warnings in the collection and clear them.
        self.assertEqual(aw.WarningSource.LAYOUT, warning_collector[0].source)
        self.assertEqual("Font '28 Days Later' has not been found. Using 'Calibri' font instead. Reason: alternative name from document.", warning_collector[0].description)
        warning_collector.clear()
        self.assertEqual(0, warning_collector.count)

    def test_substitution_warnings(self):
        doc = aw.Document(file_name=MY_DIR + 'Rendering.docx')
        callback = aw.WarningInfoCollection()
        doc.warning_callback = callback
        font_settings = aw.fonts.FontSettings()
        font_settings.substitution_settings.default_font_substitution.default_font_name = 'Arial'
        font_settings.set_fonts_folder(FONTS_DIR, False)
        font_settings.substitution_settings.table_substitution.add_substitutes('Arial', ['Arvo', 'Slab'])
        doc.font_settings = font_settings
        doc.save(file_name=ARTIFACTS_DIR + 'FontSettings.SubstitutionWarnings.pdf')
        self.assertEqual("Font 'Arial' has not been found. Using 'Arvo' font instead. Reason: table substitution.", callback[0].description)
        self.assertEqual("Font 'Times New Roman' has not been found. Using 'M+ 2m' font instead. Reason: font info substitution.", callback[1].description)

    def test_font_source_file(self):
        #ExStart
        #ExFor:FileFontSource
        #ExFor:FileFontSource.__init__(str)
        #ExFor:FileFontSource.__init__(str,int)
        #ExFor:FileFontSource.file_path
        #ExFor:FileFontSource.type
        #ExFor:FontSourceBase
        #ExFor:FontSourceBase.priority
        #ExFor:FontSourceBase.type
        #ExFor:FontSourceType
        #ExSummary:Shows how to use a font file in the local file system as a font source.
        file_font_source = aw.fonts.FileFontSource(file_path=MY_DIR + 'Alte DIN 1451 Mittelschrift.ttf', priority=0)
        doc = aw.Document()
        doc.font_settings = aw.fonts.FontSettings()
        doc.font_settings.set_fonts_sources(sources=[file_font_source])
        self.assertEqual(MY_DIR + 'Alte DIN 1451 Mittelschrift.ttf', file_font_source.file_path)
        self.assertEqual(aw.fonts.FontSourceType.FONT_FILE, file_font_source.type)
        self.assertEqual(0, file_font_source.priority)
        #ExEnd

    def test_font_source_folder(self):
        #ExStart
        #ExFor:FolderFontSource
        #ExFor:FolderFontSource.__init__(str,bool)
        #ExFor:FolderFontSource.__init__(str,bool,int)
        #ExFor:FolderFontSource.folder_path
        #ExFor:FolderFontSource.scan_subfolders
        #ExFor:FolderFontSource.type
        #ExSummary:Shows how to use a local system folder which contains fonts as a font source.
        # Create a font source from a folder that contains font files.
        folder_font_source = aw.fonts.FolderFontSource(folder_path=FONTS_DIR, scan_subfolders=False, priority=1)
        doc = aw.Document()
        doc.font_settings = aw.fonts.FontSettings()
        doc.font_settings.set_fonts_sources(sources=[folder_font_source])
        self.assertEqual(FONTS_DIR, folder_font_source.folder_path)
        self.assertEqual(False, folder_font_source.scan_subfolders)
        self.assertEqual(aw.fonts.FontSourceType.FONTS_FOLDER, folder_font_source.type)
        self.assertEqual(1, folder_font_source.priority)
        #ExEnd

    def test_set_fonts_folder(self):
        for recursive in [False, True]:
            #ExStart
            #ExFor:FontSettings
            #ExFor:FontSettings.set_fonts_folder(str,bool)
            #ExSummary:Shows how to set a font source directory.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.font.name = 'Arvo'
            builder.writeln('Hello world!')
            builder.font.name = 'Amethysta'
            builder.writeln('The quick brown fox jumps over the lazy dog.')
            # Our font sources do not contain the font that we have used for text in this document.
            # If we use these font settings while rendering this document,
            # Aspose.Words will apply a fallback font to text which has a font that Aspose.Words cannot locate.
            original_font_sources = aw.fonts.FontSettings.default_instance.get_fonts_sources()
            self.assertEqual(1, len(original_font_sources))
            self.assertTrue(any([f.full_font_name == 'Arial' for f in original_font_sources[0].get_available_fonts()]))
            # The default font sources are missing the two fonts that we are using in this document.
            self.assertFalse(any([f.full_font_name == 'Arvo' for f in original_font_sources[0].get_available_fonts()]))
            self.assertFalse(any([f.full_font_name == 'Amethysta' for f in original_font_sources[0].get_available_fonts()]))
            # Use the "SetFontsFolder" method to set a directory which will act as a new font source.
            # Pass "false" as the "recursive" argument to include fonts from all the font files that are in the directory
            # that we are passing in the first argument, but not include any fonts in any of that directory's subfolders.
            # Pass "true" as the "recursive" argument to include all font files in the directory that we are passing
            # in the first argument, as well as all the fonts in its subdirectories.
            aw.fonts.FontSettings.default_instance.set_fonts_folder(FONTS_DIR, recursive)
            new_font_sources = aw.fonts.FontSettings.default_instance.get_fonts_sources()
            self.assertEqual(1, len(new_font_sources))
            self.assertFalse(any([f.full_font_name == 'Arial' for f in new_font_sources[0].get_available_fonts()]))
            self.assertTrue(any([f.full_font_name == 'Arvo' for f in new_font_sources[0].get_available_fonts()]))
            # The "Amethysta" font is in a subfolder of the font directory.
            if recursive:
                self.assertEqual(30, len(new_font_sources[0].get_available_fonts()))
                self.assertTrue(any([f.full_font_name == 'Amethysta' for f in new_font_sources[0].get_available_fonts()]))
            else:
                self.assertEqual(18, len(new_font_sources[0].get_available_fonts()))
                self.assertFalse(any([f.full_font_name == 'Amethysta' for f in new_font_sources[0].get_available_fonts()]))
            doc.save(file_name=ARTIFACTS_DIR + 'FontSettings.SetFontsFolder.pdf')
            # Restore the original font sources.
            aw.fonts.FontSettings.default_instance.set_fonts_sources(sources=original_font_sources)
            #ExEnd

    def test_set_fonts_folders(self):
        for recursive in [False, True]:
            #ExStart
            #ExFor:FontSettings
            #ExFor:FontSettings.set_fonts_folders(List[str],bool)
            #ExSummary:Shows how to set multiple font source directories.
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc=doc)
            builder.font.name = 'Amethysta'
            builder.writeln('The quick brown fox jumps over the lazy dog.')
            builder.font.name = 'Junction Light'
            builder.writeln('The quick brown fox jumps over the lazy dog.')
            # Our font sources do not contain the font that we have used for text in this document.
            # If we use these font settings while rendering this document,
            # Aspose.Words will apply a fallback font to text which has a font that Aspose.Words cannot locate.
            original_font_sources = aw.fonts.FontSettings.default_instance.get_fonts_sources()
            self.assertEqual(1, len(original_font_sources))
            self.assertTrue(any([f.full_font_name == 'Arial' for f in original_font_sources[0].get_available_fonts()]))
            # The default font sources are missing the two fonts that we are using in this document.
            self.assertFalse(any([f.full_font_name == 'Amethysta' for f in original_font_sources[0].get_available_fonts()]))
            self.assertFalse(any([f.full_font_name == 'Junction Light' for f in original_font_sources[0].get_available_fonts()]))
            # Use the "SetFontsFolders" method to create a font source from each font directory that we pass as the first argument.
            # Pass "false" as the "recursive" argument to include fonts from all the font files that are in the directories
            # that we are passing in the first argument, but not include any fonts from any of the directories' subfolders.
            # Pass "true" as the "recursive" argument to include all font files in the directories that we are passing
            # in the first argument, as well as all the fonts in their subdirectories.
            aw.fonts.FontSettings.default_instance.set_fonts_folders([FONTS_DIR + '/Amethysta', FONTS_DIR + '/Junction'], recursive)
            new_font_sources = aw.fonts.FontSettings.default_instance.get_fonts_sources()
            self.assertEqual(2, len(new_font_sources))
            self.assertFalse(any([f.full_font_name == 'Arial' for f in new_font_sources[0].get_available_fonts()]))
            self.assertEqual(1, len(new_font_sources[0].get_available_fonts()))
            self.assertTrue(any([f.full_font_name == 'Amethysta' for f in new_font_sources[0].get_available_fonts()]))
            # The "Junction" folder itself contains no font files, but has subfolders that do.
            if recursive:
                self.assertEqual(11, len(new_font_sources[1].get_available_fonts()))
                self.assertTrue(any([f.full_font_name == 'Junction Light' for f in new_font_sources[1].get_available_fonts()]))
            else:
                self.assertEqual(0, len(new_font_sources[1].get_available_fonts()))
            doc.save(file_name=ARTIFACTS_DIR + 'FontSettings.SetFontsFolders.pdf')
            # Restore the original font sources.
            aw.fonts.FontSettings.default_instance.set_fonts_sources(sources=original_font_sources)
            #ExEnd

    def test_add_font_source(self):
        #ExStart
        #ExFor:FontSettings
        #ExFor:FontSettings.get_fonts_sources()
        #ExFor:FontSettings.set_fonts_sources(List[FontSourceBase])
        #ExSummary:Shows how to add a font source to our existing font sources.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.name = 'Arial'
        builder.writeln('Hello world!')
        builder.font.name = 'Amethysta'
        builder.writeln('The quick brown fox jumps over the lazy dog.')
        builder.font.name = 'Junction Light'
        builder.writeln('The quick brown fox jumps over the lazy dog.')
        original_font_sources = aw.fonts.FontSettings.default_instance.get_fonts_sources()
        self.assertEqual(1, len(original_font_sources))
        self.assertTrue(any([f.full_font_name == 'Arial' for f in original_font_sources[0].get_available_fonts()]))
        # The default font source is missing two of the fonts that we are using in our document.
        # When we save this document, Aspose.Words will apply fallback fonts to all text formatted with inaccessible fonts.
        self.assertFalse(any([f.full_font_name == 'Amethysta' for f in original_font_sources[0].get_available_fonts()]))
        self.assertFalse(any([f.full_font_name == 'Junction Light' for f in original_font_sources[0].get_available_fonts()]))
        # Create a font source from a folder that contains fonts.
        folder_font_source = aw.fonts.FolderFontSource(folder_path=FONTS_DIR, scan_subfolders=True)
        # Apply a new array of font sources that contains the original font sources, as well as our custom fonts.
        updated_font_sources = [original_font_sources[0], folder_font_source]
        aw.fonts.FontSettings.default_instance.set_fonts_sources(sources=updated_font_sources)
        # Verify that Aspose.Words has access to all required fonts before we render the document to PDF.
        updated_font_sources = aw.fonts.FontSettings.default_instance.get_fonts_sources()
        self.assertTrue(any([f.full_font_name == 'Arial' for f in updated_font_sources[0].get_available_fonts()]))
        self.assertTrue(any([f.full_font_name == 'Amethysta' for f in updated_font_sources[1].get_available_fonts()]))
        self.assertTrue(any([f.full_font_name == 'Junction Light' for f in updated_font_sources[1].get_available_fonts()]))
        doc.save(file_name=ARTIFACTS_DIR + 'FontSettings.AddFontSource.pdf')
        # Restore the original font sources.
        aw.fonts.FontSettings.default_instance.set_fonts_sources(sources=original_font_sources)
        #ExEnd

    def test_set_specify_font_folder(self):
        font_settings = aw.fonts.FontSettings()
        font_settings.set_fonts_folder(FONTS_DIR, False)
        # Using load options
        load_options = aw.loading.LoadOptions()
        load_options.font_settings = font_settings
        doc = aw.Document(file_name=MY_DIR + 'Rendering.docx', load_options=load_options)
        folder_source = doc.font_settings.get_fonts_sources()[0].as_folder_font_source()
        self.assertEqual(FONTS_DIR, folder_source.folder_path)
        self.assertFalse(folder_source.scan_subfolders)

    def test_table_substitution(self):
        #ExStart
        #ExFor:Document.font_settings
        #ExFor:TableSubstitutionRule.set_substitutes(str,List[str])
        #ExSummary:Shows how set font substitution rules.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.name = 'Arial'
        builder.writeln('Hello world!')
        builder.font.name = 'Amethysta'
        builder.writeln('The quick brown fox jumps over the lazy dog.')
        font_sources = aw.fonts.FontSettings.default_instance.get_fonts_sources()
        # The default font sources contain the first font that the document uses.
        self.assertEqual(1, len(font_sources))
        self.assertTrue(any([f.full_font_name == 'Arial' for f in font_sources[0].get_available_fonts()]))
        # The second font, "Amethysta", is unavailable.
        self.assertFalse(any([f.full_font_name == 'Amethysta' for f in font_sources[0].get_available_fonts()]))
        # We can configure a font substitution table which determines
        # which fonts Aspose.Words will use as substitutes for unavailable fonts.
        # Set two substitution fonts for "Amethysta": "Arvo", and "Courier New".
        # If the first substitute is unavailable, Aspose.Words attempts to use the second substitute, and so on.
        doc.font_settings = aw.fonts.FontSettings()
        doc.font_settings.substitution_settings.table_substitution.set_substitutes('Amethysta', ['Arvo', 'Courier New'])
        # "Amethysta" is unavailable, and the substitution rule states that the first font to use as a substitute is "Arvo".
        self.assertFalse(any([f.full_font_name == 'Arvo' for f in font_sources[0].get_available_fonts()]))
        # "Arvo" is also unavailable, but "Courier New" is.
        self.assertTrue(any([f.full_font_name == 'Courier New' for f in font_sources[0].get_available_fonts()]))
        # The output document will display the text that uses the "Amethysta" font formatted with "Courier New".
        doc.save(file_name=ARTIFACTS_DIR + 'FontSettings.TableSubstitution.pdf')
        #ExEnd

    def test_set_specify_font_folders(self):
        font_settings = aw.fonts.FontSettings()
        font_settings.set_fonts_folders([FONTS_DIR, 'C:\\Windows\\Fonts\\'], True)
        # Using load options
        load_options = aw.loading.LoadOptions()
        load_options.font_settings = font_settings
        doc = aw.Document(file_name=MY_DIR + 'Rendering.docx', load_options=load_options)
        folder_source = doc.font_settings.get_fonts_sources()[0].as_folder_font_source()
        self.assertEqual(FONTS_DIR, folder_source.folder_path)
        self.assertTrue(folder_source.scan_subfolders)
        folder_source = doc.font_settings.get_fonts_sources()[1].as_folder_font_source()
        self.assertEqual('C:\\Windows\\Fonts\\', folder_source.folder_path)
        self.assertTrue(folder_source.scan_subfolders)

    def test_font_source_memory(self):
        #ExStart
        #ExFor:MemoryFontSource
        #ExFor:MemoryFontSource.__init__(bytes)
        #ExFor:MemoryFontSource.__init__(bytes,int)
        #ExFor:MemoryFontSource.font_data
        #ExFor:MemoryFontSource.type
        #ExSummary:Shows how to use a byte array with data from a font file as a font source.
        font_bytes = system_helper.io.File.read_all_bytes(MY_DIR + 'Alte DIN 1451 Mittelschrift.ttf')
        memory_font_source = aw.fonts.MemoryFontSource(font_data=font_bytes, priority=0)
        doc = aw.Document()
        doc.font_settings = aw.fonts.FontSettings()
        doc.font_settings.set_fonts_sources(sources=[memory_font_source])
        self.assertEqual(aw.fonts.FontSourceType.MEMORY_FONT, memory_font_source.type)
        self.assertEqual(0, memory_font_source.priority)
        #ExEnd

    def test_load_font_fallback_settings_from_file(self):
        #ExStart
        #ExFor:FontFallbackSettings.load(str)
        #ExFor:FontFallbackSettings.save(str)
        #ExSummary:Shows how to load and save font fallback settings to/from an XML document in the local file system.
        doc = aw.Document(file_name=MY_DIR + 'Rendering.docx')
        # Load an XML document that defines a set of font fallback settings.
        font_settings = aw.fonts.FontSettings()
        font_settings.fallback_settings.load(file_name=MY_DIR + 'Font fallback rules.xml')
        doc.font_settings = font_settings
        doc.save(file_name=ARTIFACTS_DIR + 'FontSettings.LoadFontFallbackSettingsFromFile.pdf')
        # Save our document's current font fallback settings as an XML document.
        doc.font_settings.fallback_settings.save(file_name=ARTIFACTS_DIR + 'FallbackSettings.xml')
        #ExEnd

    def test_load_noto_fonts_fallback_settings(self):
        #ExStart
        #ExFor:FontFallbackSettings.load_noto_fallback_settings
        #ExSummary:Shows how to add predefined font fallback settings for Google Noto fonts.
        font_settings = aw.fonts.FontSettings()
        # These are free fonts licensed under the SIL Open Font License.
        # We can download the fonts here:
        # https:#www.google.com/get/noto/#sans-lgc
        font_settings.set_fonts_folder(FONTS_DIR + 'Noto', False)
        # Note that the predefined settings only use Sans-style Noto fonts with regular weight.
        # Some of the Noto fonts use advanced typography features.
        # Fonts featuring advanced typography may not be rendered correctly as Aspose.Words currently do not support them.
        font_settings.fallback_settings.load_noto_fallback_settings()
        font_settings.substitution_settings.font_info_substitution.enabled = False
        font_settings.substitution_settings.default_font_substitution.default_font_name = 'Noto Sans'
        doc = aw.Document()
        doc.font_settings = font_settings
        #ExEnd

    def test_default_font_substitution_rule(self):
        #ExStart
        #ExFor:DefaultFontSubstitutionRule
        #ExFor:DefaultFontSubstitutionRule.default_font_name
        #ExFor:FontSubstitutionSettings.default_font_substitution
        #ExSummary:Shows how to set the default font substitution rule.
        doc = aw.Document()
        font_settings = aw.fonts.FontSettings()
        doc.font_settings = font_settings
        # Get the default substitution rule within FontSettings.
        # This rule will substitute all missing fonts with "Times New Roman".
        default_font_substitution_rule = font_settings.substitution_settings.default_font_substitution
        self.assertTrue(default_font_substitution_rule.enabled)
        self.assertEqual('Times New Roman', default_font_substitution_rule.default_font_name)
        # Set the default font substitute to "Courier New".
        default_font_substitution_rule.default_font_name = 'Courier New'
        # Using a document builder, add some text in a font that we do not have to see the substitution take place,
        # and then render the result in a PDF.
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.name = 'Missing Font'
        builder.writeln('Line written in a missing font, which will be substituted with Courier New.')
        doc.save(file_name=ARTIFACTS_DIR + 'FontSettings.DefaultFontSubstitutionRule.pdf')
        #ExEnd
        self.assertEqual('Courier New', default_font_substitution_rule.default_font_name)

    def test_resolve_fonts_before_loading_document(self):
        #ExStart
        #ExFor:LoadOptions.font_settings
        #ExSummary:Shows how to designate font substitutes during loading.
        load_options = aw.loading.LoadOptions()
        load_options.font_settings = aw.fonts.FontSettings()
        # Set a font substitution rule for a LoadOptions object.
        # If the document we are loading uses a font which we do not have,
        # this rule will substitute the unavailable font with one that does exist.
        # In this case, all uses of the "MissingFont" will convert to "Comic Sans MS".
        substitution_rule = load_options.font_settings.substitution_settings.table_substitution
        substitution_rule.add_substitutes('MissingFont', ['Comic Sans MS'])
        doc = aw.Document(file_name=MY_DIR + 'Missing font.html', load_options=load_options)
        # At this point such text will still be in "MissingFont".
        # Font substitution will take place when we render the document.
        self.assertEqual('MissingFont', doc.first_section.body.first_paragraph.runs[0].font.name)
        doc.save(file_name=ARTIFACTS_DIR + 'FontSettings.ResolveFontsBeforeLoadingDocument.pdf')
        #ExEnd

    def test_add_font_substitutes(self):
        font_settings = aw.fonts.FontSettings()
        font_settings.substitution_settings.table_substitution.set_substitutes('Slab', ['Times New Roman', 'Arial'])
        font_settings.substitution_settings.table_substitution.add_substitutes('Arvo', ['Open Sans', 'Arial'])
        doc = aw.Document(file_name=MY_DIR + 'Rendering.docx')
        doc.font_settings = font_settings
        alternative_fonts = list(doc.font_settings.substitution_settings.table_substitution.get_substitutes('Slab'))
        self.assertSequenceEqual(['Times New Roman', 'Arial'], alternative_fonts)
        alternative_fonts = list(doc.font_settings.substitution_settings.table_substitution.get_substitutes('Arvo'))
        self.assertSequenceEqual(['Open Sans', 'Arial'], alternative_fonts)

    def test_font_source_system(self):
        #ExStart
        #ExFor:TableSubstitutionRule.add_substitutes(str,List[str])
        #ExFor:FontSubstitutionRule.enabled
        #ExFor:TableSubstitutionRule.get_substitutes(str)
        #ExFor:FontSettings.reset_font_sources
        #ExFor:FontSettings.substitution_settings
        #ExFor:FontSubstitutionSettings
        #ExFor:SystemFontSource
        #ExFor:SystemFontSource.get_system_font_folders
        #ExFor:SystemFontSource.type
        #ExSummary:Shows how to access a document's system font source and set font substitutes.
        doc = aw.Document()
        doc.font_settings = aw.fonts.FontSettings()
        # By default, a blank document always contains a system font source.
        self.assertEqual(1, len(doc.font_settings.get_fonts_sources()))
        system_font_source = doc.font_settings.get_fonts_sources()[0]
        self.assertEqual(aw.fonts.FontSourceType.SYSTEM_FONTS, system_font_source.type)
        self.assertEqual(0, system_font_source.priority)
        if platform.system() == 'Windows':
            fonts_path = 'C:\\WINDOWS\\Fonts'
            self.assertEqual(fonts_path.lower(), aw.fonts.SystemFontSource.get_system_font_folders()[0].lower())
        for system_font_folder in aw.fonts.SystemFontSource.get_system_font_folders():
            print(system_font_folder)
        # Set a font that exists in the Windows Fonts directory as a substitute for one that does not.
        doc.font_settings.substitution_settings.font_info_substitution.enabled = True
        doc.font_settings.substitution_settings.table_substitution.add_substitutes('Kreon-Regular', ['Calibri'])
        self.assertEqual(1, len(list(doc.font_settings.substitution_settings.table_substitution.get_substitutes('Kreon-Regular'))))
        self.assertIn('Calibri', list(doc.font_settings.substitution_settings.table_substitution.get_substitutes('Kreon-Regular')))
        # Alternatively, we could add a folder font source in which the corresponding folder contains the font.
        folder_font_source = aw.fonts.FolderFontSource(FONTS_DIR, False)
        doc.font_settings.set_fonts_sources([system_font_source, folder_font_source])
        self.assertEqual(2, len(doc.font_settings.get_fonts_sources()))
        # Resetting the font sources still leaves us with the system font source as well as our substitutes.
        doc.font_settings.reset_font_sources()
        self.assertEqual(1, len(doc.font_settings.get_fonts_sources()))
        self.assertEqual(aw.fonts.FontSourceType.SYSTEM_FONTS, doc.font_settings.get_fonts_sources()[0].type)
        self.assertEqual(1, len(list(doc.font_settings.substitution_settings.table_substitution.get_substitutes('Kreon-Regular'))))
        #ExEnd

    def test_load_font_fallback_settings_from_stream(self):
        #ExStart
        #ExFor:FontFallbackSettings.load(BytesIO)
        #ExFor:FontFallbackSettings.save(BytesIO)
        #ExSummary:Shows how to load and save font fallback settings to/from a stream.
        doc = aw.Document(MY_DIR + 'Rendering.docx')
        # Load an XML document that defines a set of font fallback settings.
        with open(MY_DIR + 'Font fallback rules.xml', 'rb') as font_fallback_stream:
            font_settings = aw.fonts.FontSettings()
            font_settings.fallback_settings.load(font_fallback_stream)
            doc.font_settings = font_settings
        doc.save(ARTIFACTS_DIR + 'FontSettings.load_font_fallback_settings_from_stream.pdf')
        # Use a stream to save our document's current font fallback settings as an XML document.
        with open(ARTIFACTS_DIR + 'FallbackSettings.xml', 'wb') as font_fallback_stream:
            doc.font_settings.fallback_settings.save(font_fallback_stream)
        #ExEnd
        fallback_settings_doc = ET.parse(ARTIFACTS_DIR + 'FallbackSettings.xml')
        rules = fallback_settings_doc.getroot().find('{Aspose.Words}FallbackTable').findall('{Aspose.Words}Rule')
        self.assertEqual('0B80-0BFF', rules[0].attrib['Ranges'])
        self.assertEqual('Vijaya', rules[0].attrib['FallbackFonts'])
        self.assertEqual('1F300-1F64F', rules[1].attrib['Ranges'])
        self.assertEqual('Segoe UI Emoji, Segoe UI Symbol', rules[1].attrib['FallbackFonts'])
        self.assertEqual('2000-206F, 2070-209F, 20B9', rules[2].attrib['Ranges'])
        self.assertEqual('Arial', rules[2].attrib['FallbackFonts'])
        self.assertEqual('3040-309F', rules[3].attrib['Ranges'])
        self.assertEqual('MS Gothic', rules[3].attrib['FallbackFonts'])
        self.assertEqual('Times New Roman', rules[3].attrib['BaseFonts'])
        self.assertEqual('3040-309F', rules[4].attrib['Ranges'])
        self.assertEqual('MS Mincho', rules[4].attrib['FallbackFonts'])
        self.assertEqual('Arial Unicode MS', rules[5].attrib['FallbackFonts'])

    @unittest.skip('Discrepancy in assertion between Python and .Net')
    def test_font_config_substitution(self):
        #ExStart
        #ExFor:FontConfigSubstitutionRule
        #ExFor:FontConfigSubstitutionRule.enabled
        #ExFor:FontConfigSubstitutionRule.is_font_config_available
        #ExFor:FontConfigSubstitutionRule.reset_cache
        #ExFor:FontSubstitutionRule
        #ExFor:FontSubstitutionRule.enabled
        #ExFor:FontSubstitutionSettings.font_config_substitution
        #ExSummary:Shows operating system-dependent font config substitution.
        font_settings = aw.fonts.FontSettings()
        font_config_substitution = font_settings.substitution_settings.font_config_substitution
        # The FontConfigSubstitutionRule object works differently on Windows/non-Windows platforms.
        # On Windows, it is unavailable.
        if platform.system() == 'Windows':
            self.assertFalse(font_config_substitution.enabled)
            self.assertFalse(font_config_substitution.is_font_config_available())
        else:
            # On Linux/Mac, we will have access to it, and will be able to perform operations.
            self.assertTrue(font_config_substitution.enabled)
            self.assertTrue(font_config_substitution.is_font_config_available())
            font_config_substitution.reset_cache()
        #ExEnd

    def test_fallback_settings(self):
        #ExStart
        #ExFor:FontFallbackSettings.load_ms_office_fallback_settings
        #ExFor:FontFallbackSettings.load_noto_fallback_settings
        #ExSummary:Shows how to load pre-defined fallback font settings.
        doc = aw.Document()
        font_settings = aw.fonts.FontSettings()
        doc.font_settings = font_settings
        font_fallback_settings = font_settings.fallback_settings
        # Save the default fallback font scheme to an XML document.
        # For example, one of the elements has a value of "0C00-0C7F" for Range and a corresponding "Vani" value for FallbackFonts.
        # This means that if the font some text is using does not have symbols for the 0x0C00-0x0C7F Unicode block,
        # the fallback scheme will use symbols from the "Vani" font substitute.
        font_fallback_settings.save(ARTIFACTS_DIR + 'FontSettings.fallback_settings.default.xml')
        # Below are two pre-defined font fallback schemes we can choose from.
        # 1 -  Use the default Microsoft Office scheme, which is the same one as the default:
        font_fallback_settings.load_ms_office_fallback_settings()
        font_fallback_settings.save(ARTIFACTS_DIR + 'FontSettings.fallback_settings.load_ms_office_fallback_settings.xml')
        # 2 -  Use the scheme built from Google Noto fonts:
        font_fallback_settings.load_noto_fallback_settings()
        font_fallback_settings.save(ARTIFACTS_DIR + 'FontSettings.fallback_settings.load_noto_fallback_settings.xml')
        #ExEnd
        fallback_settings_doc = ET.parse(ARTIFACTS_DIR + 'FontSettings.fallback_settings.default.xml')
        rules = fallback_settings_doc.getroot().find('{Aspose.Words}FallbackTable').findall('{Aspose.Words}Rule')
        self.assertEqual('0C00-0C7F', rules[9].attrib['Ranges'])
        self.assertEqual('Vani', rules[9].attrib['FallbackFonts'])

    def test_fallback_settings_custom(self):
        #ExStart
        #ExFor:FontSettings.fallback_settings
        #ExFor:FontFallbackSettings
        #ExFor:FontFallbackSettings.build_automatic
        #ExSummary:Shows how to distribute fallback fonts across Unicode character code ranges.
        doc = aw.Document()
        font_settings = aw.fonts.FontSettings()
        doc.font_settings = font_settings
        font_fallback_settings = font_settings.fallback_settings
        # Configure our font settings to source fonts only from the "MyFonts" folder.
        folder_font_source = aw.fonts.FolderFontSource(FONTS_DIR, False)
        font_settings.set_fonts_sources([folder_font_source])
        # Calling the "build_automatic" method will generate a fallback scheme that
        # distributes accessible fonts across as many Unicode character codes as possible.
        # In our case, it only has access to the handful of fonts inside the "MyFonts" folder.
        font_fallback_settings.build_automatic()
        font_fallback_settings.save(ARTIFACTS_DIR + 'FontSettings.fallback_settings_custom.build_automatic.xml')
        # We can also load a custom substitution scheme from a file like this.
        # This scheme applies the "AllegroOpen" font across the "0000-00ff" Unicode blocks, the "AllegroOpen" font across "0100-024f",
        # and the "M+ 2m" font in all other ranges that other fonts in the scheme do not cover.
        font_fallback_settings.load(MY_DIR + 'Custom font fallback settings.xml')
        # Create a document builder and set its font to one that does not exist in any of our sources.
        # Our font settings will invoke the fallback scheme for characters that we type using the unavailable font.
        builder = aw.DocumentBuilder(doc)
        builder.font.name = 'Missing Font'
        # Use the builder to print every Unicode character from 0x0021 to 0x052F,
        # with descriptive lines dividing Unicode blocks we defined in our custom font fallback scheme.
        for i in range(33, 1328):
            if i == 33:
                builder.writeln('\n\n0x0021 - 0x00FF: \nBasic Latin/Latin-1 Supplement Unicode blocks in "AllegroOpen" font:')
            elif i == 256:
                builder.writeln('\n\n0x0100 - 0x024F: \nLatin Extended A/B blocks, mostly in "AllegroOpen" font:')
            elif i == 592:
                builder.writeln('\n\n0x0250 - 0x052F: \nIPA/Greek/Cyrillic blocks in "M+ 2m" font:')
            builder.write(chr(i))
        doc.save(ARTIFACTS_DIR + 'FontSettings.fallback_settings_custom.pdf')
        #ExEnd
        fallback_settings_doc = ET.parse(ARTIFACTS_DIR + 'FontSettings.fallback_settings_custom.build_automatic.xml')
        rules = fallback_settings_doc.getroot().find('{Aspose.Words}FallbackTable').findall('{Aspose.Words}Rule')
        self.assertEqual('0000-007F', rules[0].attrib['Ranges'])
        self.assertEqual('AllegroOpen', rules[0].attrib['FallbackFonts'])
        self.assertEqual('0100-017F', rules[2].attrib['Ranges'])
        self.assertEqual('AllegroOpen', rules[2].attrib['FallbackFonts'])
        self.assertEqual('0250-02AF', rules[4].attrib['Ranges'])
        self.assertEqual('M+ 2m', rules[4].attrib['FallbackFonts'])
        self.assertEqual('0370-03FF', rules[7].attrib['Ranges'])
        self.assertEqual('Arvo', rules[7].attrib['FallbackFonts'])

    def test_table_substitution_rule(self):
        #ExStart
        #ExFor:TableSubstitutionRule
        #ExFor:TableSubstitutionRule.load_linux_settings
        #ExFor:TableSubstitutionRule.load_windows_settings
        #ExFor:TableSubstitutionRule.save(BytesIO)
        #ExFor:TableSubstitutionRule.save(str)
        #ExSummary:Shows how to access font substitution tables for Windows and Linux.
        doc = aw.Document()
        font_settings = aw.fonts.FontSettings()
        doc.font_settings = font_settings
        # Create a new table substitution rule and load the default Microsoft Windows font substitution table.
        table_substitution_rule = font_settings.substitution_settings.table_substitution
        table_substitution_rule.load_windows_settings()
        # In Windows, the default substitute for the "Times New Roman CE" font is "Times New Roman".
        self.assertListEqual(['Times New Roman'], list(table_substitution_rule.get_substitutes('Times New Roman CE')))
        # We can save the table in the form of an XML document.
        table_substitution_rule.save(ARTIFACTS_DIR + 'FontSettings.table_substitution_rule.windows.xml')
        # Linux has its own substitution table.
        # There are multiple substitute fonts for "Times New Roman CE".
        # If the first substitute, "FreeSerif" is also unavailable,
        # this rule will cycle through the others in the array until it finds an available one.
        table_substitution_rule.load_linux_settings()
        self.assertListEqual(['FreeSerif', 'Liberation Serif', 'DejaVu Serif'], list(table_substitution_rule.get_substitutes('Times New Roman CE')))
        # Save the Linux substitution table in the form of an XML document using a stream.
        with open(ARTIFACTS_DIR + 'FontSettings.table_substitution_rule.linux.xml', 'wb') as file_stream:
            table_substitution_rule.save(file_stream)
        #ExEnd
        fallback_settings_doc = ET.parse(ARTIFACTS_DIR + 'FontSettings.table_substitution_rule.windows.xml')
        rules = fallback_settings_doc.getroot().find('{Aspose.Words}SubstitutesTable').findall('{Aspose.Words}Item')
        self.assertEqual('Times New Roman CE', rules[16].attrib['OriginalFont'])
        self.assertEqual('Times New Roman', rules[16].attrib['SubstituteFonts'])
        fallback_settings_doc = ET.parse(ARTIFACTS_DIR + 'FontSettings.table_substitution_rule.linux.xml')
        rules = fallback_settings_doc.getroot().find('{Aspose.Words}SubstitutesTable').findall('{Aspose.Words}Item')
        self.assertEqual('Times New Roman CE', rules[31].attrib['OriginalFont'])
        self.assertEqual('FreeSerif, Liberation Serif, DejaVu Serif', rules[31].attrib['SubstituteFonts'])

    def test_table_substitution_rule_custom(self):
        #ExStart
        #ExFor:FontSubstitutionSettings.table_substitution
        #ExFor:TableSubstitutionRule.add_substitutes(str,List[str])
        #ExFor:TableSubstitutionRule.get_substitutes(str)
        #ExFor:TableSubstitutionRule.load(BytesIO)
        #ExFor:TableSubstitutionRule.load(str)
        #ExFor:TableSubstitutionRule.set_substitutes(str,List[str])
        #ExSummary:Shows how to work with custom font substitution tables.
        doc = aw.Document()
        font_settings = aw.fonts.FontSettings()
        doc.font_settings = font_settings
        # Create a new table substitution rule and load the default Windows font substitution table.
        table_substitution_rule = font_settings.substitution_settings.table_substitution
        # If we select fonts exclusively from our folder, we will need a custom substitution table.
        # We will no longer have access to the Microsoft Windows fonts,
        # such as "Arial" or "Times New Roman" since they do not exist in our new font folder.
        folder_font_source = aw.fonts.FolderFontSource(folder_path=FONTS_DIR, scan_subfolders=False)
        font_settings.set_fonts_sources(sources=[folder_font_source])
        # Below are two ways of loading a substitution table from a file in the local file system.
        # 1 -  From a stream:
        with system_helper.io.FileStream(MY_DIR + 'Font substitution rules.xml', system_helper.io.FileMode.OPEN) as file_stream:
            table_substitution_rule.load(stream=file_stream)
        # 2 -  Directly from a file:
        table_substitution_rule.load(file_name=MY_DIR + 'Font substitution rules.xml')
        # Since we no longer have access to "Arial", our font table will first try substitute it with "Nonexistent Font".
        # We do not have this font so that it will move onto the next substitute, "Kreon", found in the "MyFonts" folder.
        self.assertSequenceEqual(['Missing Font', 'Kreon'], list(table_substitution_rule.get_substitutes('Arial')))
        # We can expand this table programmatically. We will add an entry that substitutes "Times New Roman" with "Arvo"
        self.assertIsNone(table_substitution_rule.get_substitutes('Times New Roman'))
        table_substitution_rule.add_substitutes('Times New Roman', ['Arvo'])
        self.assertSequenceEqual(['Arvo'], list(table_substitution_rule.get_substitutes('Times New Roman')))
        # We can add a secondary fallback substitute for an existing font entry with AddSubstitutes().
        # In case "Arvo" is unavailable, our table will look for "M+ 2m" as a second substitute option.
        table_substitution_rule.add_substitutes('Times New Roman', ['M+ 2m'])
        self.assertSequenceEqual(['Arvo', 'M+ 2m'], list(table_substitution_rule.get_substitutes('Times New Roman')))
        # SetSubstitutes() can set a new list of substitute fonts for a font.
        table_substitution_rule.set_substitutes('Times New Roman', ['Squarish Sans CT', 'M+ 2m'])
        self.assertSequenceEqual(['Squarish Sans CT', 'M+ 2m'], list(table_substitution_rule.get_substitutes('Times New Roman')))
        # Writing text in fonts that we do not have access to will invoke our substitution rules.
        builder = aw.DocumentBuilder(doc=doc)
        builder.font.name = 'Arial'
        builder.writeln('Text written in Arial, to be substituted by Kreon.')
        builder.font.name = 'Times New Roman'
        builder.writeln('Text written in Times New Roman, to be substituted by Squarish Sans CT.')
        doc.save(file_name=ARTIFACTS_DIR + 'FontSettings.TableSubstitutionRule.Custom.pdf')
        #ExEnd