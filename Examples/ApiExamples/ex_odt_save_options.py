# -*- coding: utf-8 -*-
# Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExOdtSaveOptions(ApiExampleBase):

    def test_odt_11_schema(self):
        for export_to_odt11_specs in (False, True):
            with self.subTest(export_to_odt11_specs=export_to_odt11_specs):
                #ExStart
                #ExFor:OdtSaveOptions
                #ExFor:OdtSaveOptions.__init__()
                #ExFor:OdtSaveOptions.is_strict_schema11
                #ExSummary:Shows how to make a saved document conform to an older ODT schema.
                doc = aw.Document(MY_DIR + 'Rendering.docx')
                save_options = aw.saving.OdtSaveOptions()
                save_options.measure_unit = aw.saving.OdtSaveMeasureUnit.CENTIMETERS
                save_options.is_strict_schema11 = export_to_odt11_specs
                doc.save(ARTIFACTS_DIR + 'OdtSaveOptions.odt11_schema.odt', save_options)
                #ExEnd
                doc = aw.Document(ARTIFACTS_DIR + 'OdtSaveOptions.odt11_schema.odt')
                self.assertEqual(aw.MeasurementUnits.CENTIMETERS, doc.layout_options.revision_options.measurement_unit)
                if export_to_odt11_specs:
                    self.assertEqual(2, doc.range.form_fields.count)
                    self.assertEqual(aw.fields.FieldType.FIELD_FORM_TEXT_INPUT, doc.range.form_fields[0].type)
                    self.assertEqual(aw.fields.FieldType.FIELD_FORM_CHECK_BOX, doc.range.form_fields[1].type)
                else:
                    self.assertEqual(3, doc.range.form_fields.count)
                    self.assertEqual(aw.fields.FieldType.FIELD_FORM_TEXT_INPUT, doc.range.form_fields[0].type)
                    self.assertEqual(aw.fields.FieldType.FIELD_FORM_CHECK_BOX, doc.range.form_fields[1].type)
                    self.assertEqual(aw.fields.FieldType.FIELD_FORM_DROP_DOWN, doc.range.form_fields[2].type)

    def test_measurement_units(self):
        for odt_save_measure_unit in (aw.saving.OdtSaveMeasureUnit.CENTIMETERS, aw.saving.OdtSaveMeasureUnit.INCHES):
            with self.subTest(odt_save_measure_unit=odt_save_measure_unit):
                #ExStart
                #ExFor:OdtSaveOptions
                #ExFor:OdtSaveOptions.measure_unit
                #ExFor:OdtSaveMeasureUnit
                #ExSummary:Shows how to use different measurement units to define style parameters of a saved ODT document.
                doc = aw.Document(MY_DIR + 'Rendering.docx')
                # When we export the document to .odt, we can use an OdtSaveOptions object to modify how we save the document.
                # We can set the "measure_unit" property to "OdtSaveMeasureUnit.CENTIMETERS"
                # to define content such as style parameters using the metric system, which Open Office uses.
                # We can set the "measure_unit" property to "OdtSaveMeasureUnit.INCHES"
                # to define content such as style parameters using the imperial system, which Microsoft Word uses.
                save_options = aw.saving.OdtSaveOptions()
                save_options.measure_unit = odt_save_measure_unit
                doc.save(ARTIFACTS_DIR + 'OdtSaveOptions.measurement_units.odt', save_options)
                #ExEnd
                if odt_save_measure_unit == aw.saving.OdtSaveMeasureUnit.CENTIMETERS:
                    self.verify_doc_package_file_contains_string('<style:paragraph-properties fo:orphans="2" fo:widows="2" style:tab-stop-distance="1.27cm" />', ARTIFACTS_DIR + 'OdtSaveOptions.measurement_units.odt', 'styles.xml')
                elif odt_save_measure_unit == aw.saving.OdtSaveMeasureUnit.INCHES:
                    self.verify_doc_package_file_contains_string('<style:paragraph-properties fo:orphans="2" fo:widows="2" style:tab-stop-distance="0.5in" />', ARTIFACTS_DIR + 'OdtSaveOptions.measurement_units.odt', 'styles.xml')

    def test_encrypt(self):
        for save_format in (aw.SaveFormat.ODT, aw.SaveFormat.OTT):
            with self.subTest(save_format=save_format):
                #ExStart
                #ExFor:OdtSaveOptions.__init__(SaveFormat)
                #ExFor:OdtSaveOptions.password
                #ExFor:OdtSaveOptions.save_format
                #ExSummary:Shows how to encrypt a saved ODT/OTT document with a password, and then load it using Aspose.Words.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.writeln('Hello world!')
                # Create a new OdtSaveOptions, and pass either "SaveFormat.ODT",
                # or "SaveFormat.OTT" as the format to save the document in.
                save_options = aw.saving.OdtSaveOptions(save_format)
                save_options.password = '@sposeEncrypted_1145'
                extension_string = aw.FileFormatUtil.save_format_to_extension(save_format)
                # If we open this document with an appropriate editor,
                # it will prompt us for the password we specified in the SaveOptions object.
                doc.save(ARTIFACTS_DIR + 'OdtSaveOptions.encrypt' + extension_string, save_options)
                doc_info = aw.FileFormatUtil.detect_file_format(ARTIFACTS_DIR + 'OdtSaveOptions.encrypt' + extension_string)
                self.assertTrue(doc_info.is_encrypted)
                # If we wish to open or edit this document again using Aspose.Words,
                # we will have to provide a LoadOptions object with the correct password to the loading constructor.
                doc = aw.Document(ARTIFACTS_DIR + 'OdtSaveOptions.encrypt' + extension_string, aw.loading.LoadOptions('@sposeEncrypted_1145'))
                self.assertEqual('Hello world!', doc.get_text().strip())
                #ExEnd