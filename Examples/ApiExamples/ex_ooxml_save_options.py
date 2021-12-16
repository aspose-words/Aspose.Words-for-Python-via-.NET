# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import unittest
import io
import os
import time
from datetime import datetime, timedelta, timezone

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, IMAGE_DIR
from testutil import TestUtil

class ExOoxmlSaveOptions(ApiExampleBase):

    def test_password(self):

        #ExStart
        #ExFor:OoxmlSaveOptions.password
        #ExSummary:Shows how to create a password encrypted Office Open XML document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello world!")

        save_options = aw.saving.OoxmlSaveOptions()
        save_options.password = "MyPassword"

        doc.save(ARTIFACTS_DIR + "OoxmlSaveOptions.password.docx", save_options)

        # We will not be able to open this document with Microsoft Word or
        # Aspose.Words without providing the correct password.
        with self.assertRaises(Exception):
            doc = aw.Document(ARTIFACTS_DIR + "OoxmlSaveOptions.password.docx")

        # Open the encrypted document by passing the correct password in a LoadOptions object.
        doc = aw.Document(ARTIFACTS_DIR + "OoxmlSaveOptions.password.docx", aw.loading.LoadOptions("MyPassword"))

        self.assertEqual("Hello world!", doc.get_text().strip())
        #ExEnd

    def test_iso29500_strict(self):

        #ExStart
        #ExFor:CompatibilityOptions
        #ExFor:CompatibilityOptions.optimize_for(MsWordVersion)
        #ExFor:OoxmlSaveOptions
        #ExFor:OoxmlSaveOptions.__init__
        #ExFor:OoxmlSaveOptions.save_format
        #ExFor:OoxmlCompliance
        #ExFor:OoxmlSaveOptions.compliance
        #ExFor:ShapeMarkupLanguage
        #ExSummary:Shows how to set an OOXML compliance specification for a saved document to adhere to.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # If we configure compatibility options to comply with Microsoft Word 2003,
        # inserting an image will define its shape using VML.
        doc.compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2003)
        builder.insert_image(IMAGE_DIR + "Transparent background logo.png")

        self.assertEqual(aw.drawing.ShapeMarkupLanguage.VML, doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().markup_language)

        # The "ISO/IEC 29500:2008" OOXML standard does not support VML shapes.
        # If we set the "compliance" property of the SaveOptions object to "OoxmlCompliance.ISO29500_2008_STRICT",
        # any document we save while passing this object will have to follow that standard.
        save_options = aw.saving.OoxmlSaveOptions()
        save_options.compliance = aw.saving.OoxmlCompliance.ISO29500_2008_STRICT
        save_options.save_format = aw.SaveFormat.DOCX

        doc.save(ARTIFACTS_DIR + "OoxmlSaveOptions.iso29500_strict.docx", save_options)

        # Our saved document defines the shape using DML to adhere to the "ISO/IEC 29500:2008" OOXML standard.
        doc = aw.Document(ARTIFACTS_DIR + "OoxmlSaveOptions.iso29500_strict.docx")

        self.assertEqual(aw.drawing.ShapeMarkupLanguage.DML, doc.get_child(aw.NodeType.SHAPE, 0, True).as_shape().markup_language)
        #ExEnd

    def test_restarting_document_list(self):

        for restart_list_at_each_section in (False, True):
            with self.subTest(restart_list_at_each_section=restart_list_at_each_section):
                #ExStart
                #ExFor:List.is_restart_at_each_section
                #ExFor:OoxmlCompliance
                #ExFor:OoxmlSaveOptions.compliance
                #ExSummary:Shows how to configure a list to restart numbering at each section.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                doc.lists.add(aw.lists.ListTemplate.NUMBER_DEFAULT)

                list = doc.lists[0]
                list.is_restart_at_each_section = restart_list_at_each_section

                # The "is_restart_at_each_section" property will only be applicable when
                # the document's OOXML compliance level is to a standard that is newer than "OoxmlComplianceCore.ECMA376".
                options = aw.saving.OoxmlSaveOptions()
                options.compliance = aw.saving.OoxmlCompliance.ISO29500_2008_TRANSITIONAL

                builder.list_format.list = list

                builder.writeln("List item 1")
                builder.writeln("List item 2")
                builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
                builder.writeln("List item 3")
                builder.writeln("List item 4")

                doc.save(ARTIFACTS_DIR + "OoxmlSaveOptions.restarting_document_list.docx", options)

                doc = aw.Document(ARTIFACTS_DIR + "OoxmlSaveOptions.restarting_document_list.docx")

                self.assertEqual(restart_list_at_each_section, doc.lists[0].is_restart_at_each_section)
                #ExEnd

    def test_last_saved_time(self):

        for update_last_saved_time_property in (False, True):
            with self.subTest(update_last_saved_time_property=update_last_saved_time_property):
                #ExStart
                #ExFor:SaveOptions.update_last_saved_time_property
                #ExSummary:Shows how to determine whether to preserve the document's "Last saved time" property when saving.
                doc = aw.Document(MY_DIR + "Document.docx")

                self.assertEqual(datetime(2021, 5, 11, 6, 32, 0, tzinfo=timezone.utc),
                    doc.built_in_document_properties.last_saved_time)

                # When we save the document to an OOXML format, we can create an OoxmlSaveOptions object
                # and then pass it to the document's saving method to modify how we save the document.
                # Set the "update_last_saved_time_property" property to "True" to
                # set the output document's "Last saved time" built-in property to the current date/time.
                # Set the "update_last_saved_time_property" property to "False" to
                # preserve the original value of the input document's "Last saved time" built-in property.
                save_options = aw.saving.OoxmlSaveOptions()
                save_options.update_last_saved_time_property = update_last_saved_time_property

                doc.save(ARTIFACTS_DIR + "OoxmlSaveOptions.last_saved_time.docx", save_options)

                doc = aw.Document(ARTIFACTS_DIR + "OoxmlSaveOptions.last_saved_time.docx")
                last_saved_time_new = doc.built_in_document_properties.last_saved_time

                if update_last_saved_time_property:
                    self.assertAlmostEqual(datetime.now(timezone.utc), last_saved_time_new, delta=timedelta(days=1))
                else:
                    self.assertEqual(datetime(2021, 5, 11, 6, 32, 0, tzinfo=timezone.utc), last_saved_time_new)
                #ExEnd

    def test_keep_legacy_control_chars(self):

        for keep_legacy_control_chars in (False, True):
            with self.subTest(keep_legacy_control_chars=keep_legacy_control_chars):
                #ExStart
                #ExFor:OoxmlSaveOptions.keep_legacy_control_chars
                #ExFor:OoxmlSaveOptions.__init__(SaveFormat)
                #ExSummary:Shows how to support legacy control characters when converting to .docx.
                doc = aw.Document(MY_DIR + "Legacy control character.doc")

                # When we save the document to an OOXML format, we can create an OoxmlSaveOptions object
                # and then pass it to the document's saving method to modify how we save the document.
                # Set the "keep_legacy_control_chars" property to "True" to preserve
                # the "ShortDateTime" legacy character while saving.
                # Set the "keep_legacy_control_chars" property to "False" to remove
                # the "ShortDateTime" legacy character from the output document.
                save_options = aw.saving.OoxmlSaveOptions(aw.SaveFormat.DOCX)
                save_options.keep_legacy_control_chars = keep_legacy_control_chars

                doc.save(ARTIFACTS_DIR + "OoxmlSaveOptions.keep_legacy_control_chars.docx", save_options)

                doc = aw.Document(ARTIFACTS_DIR + "OoxmlSaveOptions.keep_legacy_control_chars.docx")

                self.assertEqual(
                    "\u0013date \\@ \"MM/dd/yyyy\"\u0014\u0015\f" if keep_legacy_control_chars else "\u001e\f",
                    doc.first_section.body.get_text())
                #ExEnd

    def test_document_compression(self):

        for compression_level in (aw.saving.CompressionLevel.MAXIMUM,
                                  aw.saving.CompressionLevel.FAST,
                                  aw.saving.CompressionLevel.NORMAL,
                                  aw.saving.CompressionLevel.SUPER_FAST):
            with self.subTest(compression_level=compression_level):
                #ExStart
                #ExFor:OoxmlSaveOptions.compression_level
                #ExFor:CompressionLevel
                #ExSummary:Shows how to specify the compression level to use while saving an OOXML document.
                doc = aw.Document(MY_DIR + "Big document.docx")

                # When we save the document to an OOXML format, we can create an OoxmlSaveOptions object
                # and then pass it to the document's saving method to modify how we save the document.
                # Set the "compression_level" property to "CompressionLevel.MAXIMUM" to apply the strongest and slowest compression.
                # Set the "compression_level" property to "CompressionLevel.NORMAL" to apply
                # the default compression that Aspose.Words uses while saving OOXML documents.
                # Set the "compression_level" property to "CompressionLevel.FAST" to apply a faster and weaker compression.
                # Set the "compression_level" property to "CompressionLevel.SUPER_FAST" to apply
                # the default compression that Microsoft Word uses.
                save_options = aw.saving.OoxmlSaveOptions(aw.SaveFormat.DOCX)
                save_options.compression_level = compression_level

                start_time = time.perf_counter()
                doc.save(ARTIFACTS_DIR + "OoxmlSaveOptions.document_compression.docx", save_options)
                elapsed_ms = 1000 * (time.perf_counter() - start_time)

                file_size = os.path.getsize(ARTIFACTS_DIR + "OoxmlSaveOptions.document_compression.docx")

                print(f"Saving operation done using the \"{compression_level}\" compression level:")
                print(f"\tDuration:\t{elapsed_ms} ms")
                print(f"\tFile Size:\t{file_size} bytes")
                #ExEnd

                if compression_level == aw.saving.CompressionLevel.MAXIMUM:
                    self.assertGreater(1266000, file_size)

                elif compression_level == aw.saving.CompressionLevel.NORMAL:
                    self.assertLess(1266900, file_size)

                elif compression_level == aw.saving.CompressionLevel.FAST:
                    self.assertLess(1269000, file_size)

                elif compression_level == aw.saving.CompressionLevel.SUPER_FAST:
                    self.assertLess(1271000, file_size)

    def test_check_file_signatures(self):

        compression_levels = [
            aw.saving.CompressionLevel.MAXIMUM,
            aw.saving.CompressionLevel.NORMAL,
            aw.saving.CompressionLevel.FAST,
            aw.saving.CompressionLevel.SUPER_FAST
        ]

        file_signatures = [
            "50 4B 03 04 14 00 02 00 08 00 ",
            "50 4B 03 04 14 00 00 00 08 00 ",
            "50 4B 03 04 14 00 04 00 08 00 ",
            "50 4B 03 04 14 00 06 00 08 00 "
        ]

        doc = aw.Document()
        save_options = aw.saving.OoxmlSaveOptions(aw.SaveFormat.DOCX)

        prev_file_size = 0
        for i in range(len(file_signatures)):

            save_options.compression_level = compression_levels[i]
            doc.save(ARTIFACTS_DIR + "OoxmlSaveOptions.check_file_signatures.docx", save_options)

            with io.BytesIO() as stream:
                with open(ARTIFACTS_DIR + "OoxmlSaveOptions.check_file_signatures.docx", "rb") as output_file_stream:

                    file_size = os.path.getsize(ARTIFACTS_DIR + "OoxmlSaveOptions.check_file_signatures.docx")
                    self.assertLess(prev_file_size, file_size)

                    TestUtil.copy_stream(output_file_stream, stream)
                    self.assertEqual(file_signatures[i], TestUtil.dump_array(bytes(stream.getvalue()), 0, 10))

                    prev_file_size = file_size

    def test_export_generator_name(self):

        #ExStart
        #ExFor:SaveOptions.export_generator_name
        #ExSummary:Shows how to disable adding name and version of Aspose.Words into produced files.
        doc = aw.Document()

        # Use https://docs.aspose.com/words/net/generator-or-producer-name-included-in-output-documents/ to know how to check the result.
        save_options = aw.saving.OoxmlSaveOptions()
        save_options.export_generator_name = False

        doc.save(ARTIFACTS_DIR + "OoxmlSaveOptions.export_generator_name.docx", save_options)
        #ExEnd
