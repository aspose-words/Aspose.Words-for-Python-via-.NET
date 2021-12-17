# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import unittest
import io

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExWordML2003SaveOptions(ApiExampleBase):

    def test_pretty_format(self):

        for pretty_format in (False, True):
            with self.subTest(pretty_format=pretty_format):
                #ExStart
                #ExFor:WordML2003SaveOptions
                #ExFor:WordML2003SaveOptions.save_format
                #ExSummary:Shows how to manage output document's raw content.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.writeln("Hello world!")

                # Create a "WordML2003SaveOptions" object to pass to the document's "save" method
                # to modify how we save the document to the WordML save format.
                options = aw.saving.WordML2003SaveOptions()

                self.assertEqual(aw.SaveFormat.WORD_ML, options.save_format)

                # Set the "pretty_format" property to "True" to apply tab character indentation and
                # newlines to make the output document's raw content easier to read.
                # Set the "pretty_format" property to "False" to save the document's raw content in one continuous body of the text.
                options.pretty_format = pretty_format

                doc.save(ARTIFACTS_DIR + "WordML2003SaveOptions.pretty_format.xml", options)

                with open(ARTIFACTS_DIR + "WordML2003SaveOptions.pretty_format.xml", "rb") as file:
                    file_contents = file.read().decode('utf-8')

                if pretty_format:
                    expected = (
                        "<o:DocumentProperties>\r\n\t\t" +
                            "<o:Revision>1</o:Revision>\r\n\t\t" +
                            "<o:TotalTime>0</o:TotalTime>\r\n\t\t" +
                            "<o:Pages>1</o:Pages>\r\n\t\t" +
                            "<o:Words>0</o:Words>\r\n\t\t" +
                            "<o:Characters>0</o:Characters>\r\n\t\t" +
                            "<o:Lines>1</o:Lines>\r\n\t\t" +
                            "<o:Paragraphs>1</o:Paragraphs>\r\n\t\t" +
                            "<o:CharactersWithSpaces>0</o:CharactersWithSpaces>\r\n\t\t" +
                            "<o:Version>11.5606</o:Version>\r\n\t" +
                        "</o:DocumentProperties>")
                    self.assertTrue(expected in file_contents)
                else:
                    expected = (
                        "<o:DocumentProperties><o:Revision>1</o:Revision><o:TotalTime>0</o:TotalTime><o:Pages>1</o:Pages>" +
                        "<o:Words>0</o:Words><o:Characters>0</o:Characters><o:Lines>1</o:Lines><o:Paragraphs>1</o:Paragraphs>" +
                        "<o:CharactersWithSpaces>0</o:CharactersWithSpaces><o:Version>11.5606</o:Version></o:DocumentProperties>")
                    self.assertTrue(expected in file_contents)
                #ExEnd

    def test_memory_optimization(self):

        for memory_optimization in (False, True):
            with self.subTest(memory_optimization=memory_optimization):
                #ExStart
                #ExFor:WordML2003SaveOptions
                #ExSummary:Shows how to manage memory optimization.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)
                builder.writeln("Hello world!")

                # Create a "WordML2003SaveOptions" object to pass to the document's "save" method
                # to modify how we save the document to the WordML save format.
                options = aw.saving.WordML2003SaveOptions()

                # Set the "memory_optimization" flag to "True" to decrease memory consumption
                # during the document's saving operation at the cost of a longer saving time.
                # Set the "memory_optimization" flag to "False" to save the document normally.
                options.memory_optimization = memory_optimization

                doc.save(ARTIFACTS_DIR + "WordML2003SaveOptions.memory_optimization.xml", options)
                #ExEnd
