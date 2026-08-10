from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, GOLDS_DIR, TEMP_DIR, IMAGE_DIR, FONTS_DIR
from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR
# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
import aspose.words.drawing
import aspose.words.fonts
import aspose.words.loading
import aspose.words.settings
import datetime
import system_helper
import test_util
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, FONTS_DIR, IMAGE_DIR, MY_DIR
import sys
import typing
import os
import api_example_base
from typing import List

class ExLoadOptions(ApiExampleBase):

    def test_convert_shape_to_office_math(self):
        for is_convert_shape_to_office_math in [True, False]:
            #ExStart
            #ExFor:LoadOptions.convert_shape_to_office_math
            #ExSummary:Shows how to convert EquationXML shapes to Office Math objects.
            load_options = aw.loading.LoadOptions()
            # Use this flag to specify whether to convert the shapes with EquationXML attributes
            # to Office Math objects and then load the document.
            load_options.convert_shape_to_office_math = is_convert_shape_to_office_math
            doc = aw.Document(file_name=MY_DIR + 'Math shapes.docx', load_options=load_options)
            if is_convert_shape_to_office_math:
                self.assertEqual(16, doc.get_child_nodes(aw.NodeType.SHAPE, True).count)
                self.assertEqual(34, doc.get_child_nodes(aw.NodeType.OFFICE_MATH, True).count)
            else:
                self.assertEqual(24, doc.get_child_nodes(aw.NodeType.SHAPE, True).count)
                self.assertEqual(0, doc.get_child_nodes(aw.NodeType.OFFICE_MATH, True).count)
            #ExEnd

    def test_set_encoding(self):
        #ExStart
        #ExFor:LoadOptions.encoding
        #ExSummary:Shows how to set the encoding with which to open a document.
        load_options = aw.loading.LoadOptions()
        load_options.encoding = system_helper.text.Encoding.ascii()
        # Load the document while passing the LoadOptions object, then verify the document's contents.
        doc = aw.Document(file_name=MY_DIR + 'English text.txt', load_options=load_options)
        self.assertTrue('This is a sample text in English.' in doc.to_string(save_format=aw.SaveFormat.TEXT))
        #ExEnd

    def test_font_settings(self):
        #ExStart
        #ExFor:LoadOptions.font_settings
        #ExSummary:Shows how to apply font substitution settings while loading a document.
        # Create a FontSettings object that will substitute the "Times New Roman" font
        # with the font "Arvo" from our "MyFonts" folder.
        font_settings = aw.fonts.FontSettings()
        font_settings.set_fonts_folder(FONTS_DIR, False)
        font_settings.substitution_settings.table_substitution.add_substitutes('Times New Roman', ['Arvo'])
        # Set that FontSettings object as a property of a newly created LoadOptions object.
        load_options = aw.loading.LoadOptions()
        load_options.font_settings = font_settings
        # Load the document, then render it as a PDF with the font substitution.
        doc = aw.Document(file_name=MY_DIR + 'Document.docx', load_options=load_options)
        doc.save(file_name=ARTIFACTS_DIR + 'LoadOptions.FontSettings.pdf')
        #ExEnd

    def test_load_options_msw_version(self):
        #ExStart
        #ExFor:LoadOptions.msw_version
        #ExSummary:Shows how to emulate the loading procedure of a specific Microsoft Word version during document loading.
        # By default, Aspose.Words load documents according to Microsoft Word 2019 specification.
        load_options = aw.loading.LoadOptions()
        self.assertEqual(aw.settings.MsWordVersion.WORD2019, load_options.msw_version)
        # This document is missing the default paragraph formatting style.
        # This default style will be regenerated when we load the document either with Microsoft Word or Aspose.Words.
        load_options.msw_version = aw.settings.MsWordVersion.WORD2007
        doc = aw.Document(file_name=MY_DIR + 'Document.docx', load_options=load_options)
        # The style's line spacing will have this value when loaded by Microsoft Word 2007 specification.
        self.assertAlmostEqual(12.95, doc.styles.default_paragraph_format.line_spacing, delta=0.01)
        #ExEnd

    @staticmethod
    def _test_load_options_warning_callback(warnings):
        self.assertEqual(aw.WarningType.MINOR_FORMATTING_LOSS, warnings[0].warning_type)
        self.assertEqual(aw.WarningSource.DOCX, warnings[0].source)
        self.assertEqual("Import of element 'shapedefaults' is not supported in Docx format by Aspose.Words.", warnings[0].description)
        self.assertEqual(aw.WarningType.MINOR_FORMATTING_LOSS, warnings[1].warning_type)
        self.assertEqual(aw.WarningSource.DOCX, warnings[1].source)
        self.assertEqual("Import of element 'extraClrSchemeLst' is not supported in Docx format by Aspose.Words.", warnings[1].description)

    def test_temp_folder(self):
        #ExStart
        #ExFor:LoadOptions.temp_folder
        #ExSummary:Shows how to use the hard drive instead of memory when loading a document.
        # When we load a document, various elements are temporarily stored in memory as the save operation occurs.
        # We can use this option to use a temporary folder in the local file system instead,
        # which will reduce our application's memory overhead.
        options = aw.loading.LoadOptions()
        options.temp_folder = ARTIFACTS_DIR + 'TempFiles'
        # The specified temporary folder must exist in the local file system before the load operation.
        system_helper.io.Directory.create_directory(options.temp_folder)
        doc = aw.Document(file_name=MY_DIR + 'Document.docx', load_options=options)
        # The folder will persist with no residual contents from the load operation.
        self.assertEqual(0, len(system_helper.io.Directory.get_files(options.temp_folder)))
        #ExEnd

    def test_set_editing_language_as_default(self):
        #ExStart
        #ExFor:LanguagePreferences.default_editing_language
        #ExSummary:Shows how set a default language when loading a document.
        load_options = aw.loading.LoadOptions()
        load_options.language_preferences.default_editing_language = aw.loading.EditingLanguage.RUSSIAN
        doc = aw.Document(file_name=MY_DIR + 'No default editing language.docx', load_options=load_options)
        locale_id = doc.styles.default_font.locale_id
        print('The document either has no any language set in defaults or it was set to Russian originally.' if locale_id == int(aw.loading.EditingLanguage.RUSSIAN) else 'The document default language was set to another than Russian language originally, so it is not overridden.')
        #ExEnd
        assert doc.styles.default_font.locale_id == int(aw.loading.EditingLanguage.RUSSIAN)
        doc = aw.Document(file_name=MY_DIR + 'No default editing language.docx')
        assert doc.styles.default_font.locale_id == int(aw.loading.EditingLanguage.ENGLISH_US)

    def test_convert_metafiles_to_png(self):
        #ExStart
        #ExFor:LoadOptions.convert_metafiles_to_png
        #ExSummary:Shows how to convert WMF/EMF to PNG during loading document.
        doc = aw.Document()
        shape = aw.drawing.Shape(doc, aw.drawing.ShapeType.IMAGE)
        shape.image_data.set_image(file_name=IMAGE_DIR + 'Windows MetaFile.wmf')
        shape.width = 100
        shape.height = 100
        doc.first_section.body.first_paragraph.append_child(shape)
        doc.save(file_name=ARTIFACTS_DIR + 'Image.CreateImageDirectly.docx')
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        test_util.TestUtil.verify_image_in_shape(1600, 1600, aw.drawing.ImageType.WMF, shape)
        load_options = aw.loading.LoadOptions()
        load_options.convert_metafiles_to_png = True
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'Image.CreateImageDirectly.docx', load_options=load_options)
        shape = doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape()
        test_util.TestUtil.verify_image_in_shape(1666, 1666, aw.drawing.ImageType.PNG, shape)
        #ExEnd

    def test_ignore_ole_data(self):
        #ExStart
        #ExFor:LoadOptions.ignore_ole_data
        #ExSummary:Shows how to ingore OLE data while loading.
        # Ignoring OLE data may reduce memory consumption and increase performance
        # without data lost in a case when destination format does not support OLE objects.
        load_options = aw.loading.LoadOptions()
        load_options.ignore_ole_data = True
        doc = aw.Document(file_name=MY_DIR + 'OLE objects.docx', load_options=load_options)
        doc.save(file_name=ARTIFACTS_DIR + 'LoadOptions.IgnoreOleData.docx')
        #ExEnd

    def test_recovery_mode(self):
        #ExStart:RecoveryMode
        #ExFor:LoadOptions.recovery_mode
        #ExFor:DocumentRecoveryMode
        #ExSummary:Shows how to try to recover a document if errors occurred during loading.
        load_options = aw.loading.LoadOptions()
        load_options.recovery_mode = aw.loading.DocumentRecoveryMode.TRY_RECOVER
        doc = aw.Document(file_name=MY_DIR + 'Corrupted footnotes.docx', load_options=load_options)
        #ExEnd:RecoveryMode
    #ExStart
    #ExFor:LoadOptions.resource_loading_callback
    #ExSummary:Shows how to handle external resources when loading Html documents (HtmlLinkedResourceLoadingCallback).

    class HtmlLinkedResourceLoadingCallback(aw.loading.IResourceLoadingCallback):

        def resource_loading(self, args):
            switch_condition = args.resource_type
            if switch_condition == aw.loading.ResourceType.CSS_STYLE_SHEET:
                print(f'External CSS Stylesheet found upon loading: {args.original_uri}')
                return aw.loading.ResourceLoadingAction.DEFAULT
            elif switch_condition == aw.loading.ResourceType.IMAGE:
                print(f'External Image found upon loading: {args.original_uri}')
                new_image_filename = 'Logo.jpg'
                print(f'\tImage will be substituted with: {new_image_filename}')
                with open(Path(IMAGE_DIR) / new_image_filename, 'rb') as f:
                    image_bytes = f.read()
                args.set_data(image_bytes)
                return aw.loading.ResourceLoadingAction.USER_PROVIDED
            return aw.loading.ResourceLoadingAction.DEFAULT
    #ExEnd
    #ExStart
    #ExFor:LoadOptions.warning_callback
    #ExSummary:Shows how to print and store warnings that occur during document loading (DocumentLoadingWarningCallback).

    class DocumentLoadingWarningCallback(aw.IWarningCallback):

        def __init__(self):
            self.m_warnings = []

        def warning(self, info):
            print(f'Warning: {info.warning_type}')
            print(f'\tSource: {info.source}')
            print(f'\tDescription: {info.description}')
            self.m_warnings.append(info)

        def get_warnings(self):
            return self.m_warnings
    #ExEnd
    #ExStart
    #ExFor:LoadOptions.progress_callback
    #ExFor:IDocumentLoadingCallback
    #ExFor:IDocumentLoadingCallback.notify
    #ExFor:DocumentLoadingArgs
    #ExFor:DocumentLoadingArgs.estimated_progress
    #ExSummary:Shows how to notify the user if document loading exceeded expected loading time (LoadingProgressCallback).

    class LoadingProgressCallback(aw.loading.IDocumentLoadingCallback):

        def __init__(self):
            self.max_duration = 0.5
            self.m_loading_started_at = datetime.datetime.now()

        def notify(self, args):
            canceled_at = datetime.datetime.now()
            elapsed_seconds = (canceled_at - m_loading_started_at).total_seconds
            if elapsed_seconds > self.max_duration:
                raise Exception()
    #ExEnd

    def test_open_chm_file(self):
        info = aw.FileFormatUtil.detect_file_format(MY_DIR + 'HTML help.chm')
        self.assertEqual(info.load_format, aw.LoadFormat.CHM)
        load_options = aw.loading.LoadOptions()
        load_options.encoding = 'windows-1251'
        doc = aw.Document(MY_DIR + 'HTML help.chm', load_options)