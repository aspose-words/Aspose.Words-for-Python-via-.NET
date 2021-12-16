import unittest
import io
from datetime import datetime
from typing import List

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

class ExRange(ApiExampleBase):

    def test_replace(self):

        #ExStart
        #ExFor:Range.replace(str,str)
        #ExSummary:Shows how to perform a find-and-replace text operation on the contents of a document.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Greetings, _FullName_!")

        # Perform a find-and-replace operation on our document's contents and verify the number of replacements that took place.
        replacement_count = doc.range.replace("_FullName_", "John Doe")

        self.assertEqual(1, replacement_count)
        self.assertEqual("Greetings, John Doe!", doc.get_text().strip())
        #ExEnd

    def test_replace_match_case(self):

        for match_case in (False, True):
            with self.subTest(match_case=match_case):
                #ExStart
                #ExFor:Range.replace(str,str,FindReplaceOptions)
                #ExFor:FindReplaceOptions
                #ExFor:FindReplaceOptions.match_case
                #ExSummary:Shows how to toggle case sensitivity when performing a find-and-replace operation.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.writeln("Ruby bought a ruby necklace.")

                # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
                options = aw.replacing.FindReplaceOptions()

                # Set the "match_case" flag to "True" to apply case sensitivity while finding strings to replace.
                # Set the "match_case" flag to "False" to ignore character case while searching for text to replace.
                options.match_case = match_case

                doc.range.replace("Ruby", "Jade", options)

                self.assertEqual(
                    "Jade bought a ruby necklace." if match_case else "Jade bought a Jade necklace.",
                    doc.get_text().strip())
                #ExEnd

    def test_replace_find_whole_words_only(self):

        for find_whole_words_only in (False, True):
            with self.subTest(find_whole_words_on=find_whole_words_only):
                #ExStart
                #ExFor:Range.replace(str,str,FindReplaceOptions)
                #ExFor:FindReplaceOptions
                #ExFor:FindReplaceOptions.find_whole_words_only
                #ExSummary:Shows how to toggle standalone word-only find-and-replace operations.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.writeln("Jackson will meet you in Jacksonville.")

                # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
                options = aw.replacing.FindReplaceOptions()

                # Set the "find_whole_words_only" flag to "True" to replace the found text if it is not a part of another word.
                # Set the "find_whole_words_only" flag to "False" to replace all text regardless of its surroundings.
                options.find_whole_words_only = find_whole_words_only

                doc.range.replace("Jackson", "Louis", options)

                self.assertEqual(
                    "Louis will meet you in Jacksonville." if find_whole_words_only else "Louis will meet you in Louisville.",
                    doc.get_text().strip())
                #ExEnd

    def test_ignore_deleted(self):

        for ignore_text_inside_delete_revisions in (False, True):
            with self.subTest(ignore_text_inside_delete_revisions=ignore_text_inside_delete_revisions):
                #ExStart
                #ExFor:FindReplaceOptions.ignore_deleted
                #ExSummary:Shows how to include or ignore text inside delete revisions during a find-and-replace operation.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.writeln("Hello world!")
                builder.writeln("Hello again!")

                # Start tracking revisions and remove the second paragraph, which will create a delete revision.
                # That paragraph will persist in the document until we accept the delete revision.
                doc.start_track_revisions("John Doe", datetime.now())
                doc.first_section.body.paragraphs[1].remove()
                doc.stop_track_revisions()

                self.assertTrue(doc.first_section.body.paragraphs[1].is_delete_revision)

                # We can use a "FindReplaceOptions" object to modify the find and replace process.
                options = aw.replacing.FindReplaceOptions()

                # Set the "ignore_deleted" flag to "True" to get the find-and-replace
                # operation to ignore paragraphs that are delete revisions.
                # Set the "ignore_deleted" flag to "False" to get the find-and-replace
                # operation to also search for text inside delete revisions.
                options.ignore_deleted = ignore_text_inside_delete_revisions

                doc.range.replace("Hello", "Greetings", options)

                self.assertEqual(
                    "Greetings world!\rHello again!" if ignore_text_inside_delete_revisions else "Greetings world!\rGreetings again!",
                    doc.get_text().strip())
                #ExEnd

    def test_ignore_inserted(self):

        for ignore_text_inside_insert_revisions in (True, False):
            with self.subTest(ignore_text_inside_insert_revisions=ignore_text_inside_insert_revisions):
                #ExStart
                #ExFor:FindReplaceOptions.ignore_inserted
                #ExSummary:Shows how to include or ignore text inside insert revisions during a find-and-replace operation.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.writeln("Hello world!")

                # Start tracking revisions and insert a paragraph. That paragraph will be an insert revision.
                doc.start_track_revisions("John Doe", datetime.now())
                builder.writeln("Hello again!")
                doc.stop_track_revisions()

                self.assertTrue(doc.first_section.body.paragraphs[1].is_insert_revision)

                # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
                options = aw.replacing.FindReplaceOptions()

                # Set the "ignore_inserted" flag to "True" to get the find-and-replace
                # operation to ignore paragraphs that are insert revisions.
                # Set the "ignore_inserted" flag to "False" to get the find-and-replace
                # operation to also search for text inside insert revisions.
                options.ignore_inserted = ignore_text_inside_insert_revisions

                doc.range.replace("Hello", "Greetings", options)

                self.assertEqual(
                    "Greetings world!\rHello again!" if ignore_text_inside_insert_revisions else "Greetings world!\rGreetings again!",
                    doc.get_text().strip())
                #ExEnd

    def test_ignore_fields(self):

        for ignore_text_inside_fields in (True, False):
            with self.subTest(ignore_text_inside_fields=ignore_text_inside_fields):
                #ExStart
                #ExFor:FindReplaceOptions.ignore_fields
                #ExSummary:Shows how to ignore text inside fields.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.writeln("Hello world!")
                builder.insert_field("QUOTE", "Hello again!")

                # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
                options = aw.replacing.FindReplaceOptions()

                # Set the "ignore_fields" flag to "True" to get the find-and-replace
                # operation to ignore text inside fields.
                # Set the "ignore_fields" flag to "False" to get the find-and-replace
                # operation to also search for text inside fields.
                options.ignore_fields = ignore_text_inside_fields

                doc.range.replace("Hello", "Greetings", options)

                if ignore_text_inside_fields:
                    self.assertEqual(
                        "Greetings world!\r\u0013QUOTE\u0014Hello again!\u0015",
                        doc.get_text().strip())
                else:
                    self.assertEqual(
                        "Greetings world!\r\u0013QUOTE\u0014Greetings again!\u0015",
                        doc.get_text().strip())
                #ExEnd

    def test_ignore_footnote(self):

        for is_ignore_footnotes in (True, False):
            with self.subTest(is_ignore_footnotes=is_ignore_footnotes):
                #ExStart
                #ExFor:FindReplaceOptions.ignore_footnotes
                #ExSummary:Shows how to ignore footnotes during a find-and-replace operation.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit.")
                builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, "Lorem ipsum dolor sit amet, consectetur adipiscing elit.")

                builder.insert_paragraph()

                builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit.")
                builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, "Lorem ipsum dolor sit amet, consectetur adipiscing elit.")

                # Set the "ignore_footnotes" flag to "True" to get the find-and-replace
                # operation to ignore text inside footnotes.
                # Set the "ignore_footnotes" flag to "False" to get the find-and-replace
                # operation to also search for text inside footnotes.
                options = aw.replacing.FindReplaceOptions()
                options.ignore_footnotes = is_ignore_footnotes
                doc.range.replace("Lorem ipsum", "Replaced Lorem ipsum", options)
                #ExEnd

                paragraphs = doc.first_section.body.paragraphs

                for para in paragraphs:
                    para = para.as_paragraph()
                    self.assertEqual("Replaced Lorem ipsum", para.runs[0].text)

                footnotes = [node.as_footnote() for node in doc.get_child_nodes(aw.NodeType.FOOTNOTE, True)]
                if is_ignore_footnotes:
                    expected_text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit."
                else:
                    expected_text = "Replaced Lorem ipsum dolor sit amet, consectetur adipiscing elit."
                
                self.assertEqual(expected_text, footnotes[0].to_string(aw.SaveFormat.TEXT).strip())
                self.assertEqual(expected_text, footnotes[1].to_string(aw.SaveFormat.TEXT).strip())

    def test_update_fields_in_range(self):

        #ExStart
        #ExFor:Range.update_fields
        #ExSummary:Shows how to update all the fields in a range.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.insert_field(" DOCPROPERTY Category")
        builder.insert_break(aw.BreakType.SECTION_BREAK_EVEN_PAGE)
        builder.insert_field(" DOCPROPERTY Category")

        # The above DOCPROPERTY fields will display the value of this built-in document property.
        doc.built_in_document_properties.category = "MyCategory"

        # If we update the value of a document property, we will need to update all the DOCPROPERTY fields to display it.
        self.assertEqual("", doc.range.fields[0].result)
        self.assertEqual("", doc.range.fields[1].result)

        # Update all the fields that are in the range of the first section.
        doc.first_section.range.update_fields()

        self.assertEqual("MyCategory", doc.range.fields[0].result)
        self.assertEqual("", doc.range.fields[1].result)
        #ExEnd

    def test_replace_with_string(self):

        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("This one is sad.")
        builder.writeln("That one is mad.")

        options = aw.replacing.FindReplaceOptions()
        options.match_case = False
        options.find_whole_words_only = True

        doc.range.replace("sad", "bad", options)

        doc.save(ARTIFACTS_DIR + "Range.replace_with_string.docx")

    def test_replace_with_regex(self):

        #ExStart
        #ExFor:Range.replace(Regex,str)
        #ExSummary:Shows how to replace all occurrences of a regular expression pattern with other text.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("I decided to get the curtains in gray, ideal for the grey-accented room.")

        doc.range.replace_regex("gr(a|e)y", "lavender")

        self.assertEqual("I decided to get the curtains in lavender, ideal for the lavender-accented room.", doc.get_text().strip())
        #ExEnd

    ##ExStart
    ##ExFor:FindReplaceOptions.replacing_callback
    ##ExFor:Range.replace(Regex,str,FindReplaceOptions)
    ##ExFor:ReplacingArgs.replacement
    ##ExFor:IReplacingCallback
    ##ExFor:IReplacingCallback.replacing
    ##ExFor:ReplacingArgs
    ##ExSummary:Shows how to replace all occurrences of a regular expression pattern with another string, while tracking all such replacements.
    #def test_replace_with_callback(self):

    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    builder.writeln("Our new location in New York City is opening tomorrow. " +
    #                    "Hope to see all our NYC-based customers at the opening!")

    #    # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    #    options = aw.replacing.FindReplaceOptions()

    #    # Set a callback that tracks any replacements that the "Replace" method will make.
    #    logger = ExRange.TextFindAndReplacementLogger()
    #    options.replacing_callback = logger

    #    doc.range.replace_regex("New York City|NYC", "Washington", options)

    #    self.assertEqual("Our new location in (Old value:\"New York City\") Washington is opening tomorrow. " +
    #                     "Hope to see all our (Old value:\"NYC\") Washington-based customers at the opening!", doc.get_text().strip())

    #    self.assertEqual("\"New York City\" converted to \"Washington\" 20 characters into a Run node.\r\n" +
    #                     "\"NYC\" converted to \"Washington\" 42 characters into a Run node.", logger.get_log().strip())

    #class TextFindAndReplacementLogger(aw.replacing.IReplacingCallback):
    #    """Maintains a log of every text replacement done by a find-and-replace operation
    #    and notes the original matched text's value."""

    #    def __init__(self):
    #        self.log = io.StringIO()

    #    def replacing(self, args: aw.replacing.ReplacingArgs) -> aw.replacing.ReplaceAction:

    #        self.log.write(f"\"{args.match.value}\" converted to \"{args.replacement}\" " +
    #                       f"{args.match_offset} characters into a {args.match_node.node_type} node.\n")

    #        args.replacement = f"(Old value:\"{args.match.value}\") {args.replacement}"
    #        return aw.replcaing.ReplaceAction.REPLACE

    #    def get_log(self) -> str:
    #        return self.log.getvalue()

    ##ExEnd

    ##ExStart
    ##ExFor:FindReplaceOptions.apply_font
    ##ExFor:FindReplaceOptions.replacing_callback
    ##ExFor:ReplacingArgs.group_index
    ##ExFor:ReplacingArgs.group_name
    ##ExFor:ReplacingArgs.match
    ##ExFor:ReplacingArgs.match_offset
    ##ExSummary:Shows how to apply a different font to new content via FindReplaceOptions.
    #def test_convert_numbers_to_hexadecimal(self):

    #    doc = aw.Document()
    #    builder = aw.DocumentBuilder(doc)

    #    builder.font.name = "Arial"
    #    builder.writeln("Numbers that the find-and-replace operation will convert to hexadecimal and highlight:\n" +
    #                    "123, 456, 789 and 17379.")

    #    # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    #    options = aw.replacing.FindReplaceOptions()

    #    # Set the "HighlightColor" property to a background color that we want to apply to the operation's resulting text.
    #    options.apply_font.highlight_color = drawing.Color.light_gray

    #    number_hexer = ExRange.NumberHexer()
    #    options.replacing_callback = number_hexer

    #    replacement_count = doc.range.replace_regex("[0-9]+", "", options)

    #    print(number_hexer.get_log())

    #    self.assertEqual(4, replacement_count)
    #    self.assertEqual("Numbers that the find-and-replace operation will convert to hexadecimal and highlight:\r" +
    #                    "0x7B, 0x1C8, 0x315 and 0x43E3.", doc.get_text().strip())
    #    self.assertEqual(4, len([node for node in doc.get_child_nodes(aw.NodeType.RUN, True) 
    #                             if node.as_run().font.highlight_color.to_argb() == drawing.Color.light_gray.to_argb()]))

    #class NumberHexer(aw.replacing.IReplacingCallback):
    #    """Replaces numeric find-and-replacement matches with their hexadecimal equivalents.
    #    Maintains a log of every replacement."""

    #    def __init__(self):
    #        self.current_replacement_number = 0
    #        self.log = io.StringIO()

    #    def replacing(self, args: aw.replacing.ReplacingArgs) -> aw.replacing.ReplaceAction:

    #        self.current_replacement_number += 1

    #        number = int(args.match.value)

    #        args.replacement = f"0x{number:X}"

    #        self.log.write(f"Match #{self.current_replacement_number}\n")
    #        self.log.write(f"\tOriginal value:\t{args.match.value}\n")
    #        self.log.write(f"\tReplacement:\t{args.replacement}\n")
    #        self.log.write(f"\tOffset in parent {args.match_node.node_type} node:\t{args.match_offset}\n")

    #        self.log.append_line(f"\tGroup index:\t{args.group_index}" if not args.group_name else f"\tGroup name:\t{args.group_name}")

    #        return aw.replacing.ReplaceAction.REPLACE

    #    def get_log(self):
    #        return self.log.getvalue()

    ##ExEnd

    def test_apply_paragraph_format(self):

        #ExStart
        #ExFor:FindReplaceOptions.apply_paragraph_format
        #ExFor:Range.replace(str,str)
        #ExSummary:Shows how to add formatting to paragraphs in which a find-and-replace operation has found matches.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Every paragraph that ends with a full stop like this one will be right aligned.")
        builder.writeln("This one will not!")
        builder.write("This one also will.")

        paragraphs = doc.first_section.body.paragraphs

        self.assertEqual(aw.ParagraphAlignment.LEFT, paragraphs[0].paragraph_format.alignment)
        self.assertEqual(aw.ParagraphAlignment.LEFT, paragraphs[1].paragraph_format.alignment)
        self.assertEqual(aw.ParagraphAlignment.LEFT, paragraphs[2].paragraph_format.alignment)

        # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        options = aw.replacing.FindReplaceOptions()

        # Set the "alignment" property to "ParagraphAlignment.Right" to right-align every paragraph
        # that contains a match that the find-and-replace operation finds.
        options.apply_paragraph_format.alignment = aw.ParagraphAlignment.RIGHT

        # Replace every full stop that is right before a paragraph break with an exclamation point.
        count = doc.range.replace(".&p", "!&p", options)

        self.assertEqual(2, count)
        self.assertEqual(aw.ParagraphAlignment.RIGHT, paragraphs[0].paragraph_format.alignment)
        self.assertEqual(aw.ParagraphAlignment.LEFT, paragraphs[1].paragraph_format.alignment)
        self.assertEqual(aw.ParagraphAlignment.RIGHT, paragraphs[2].paragraph_format.alignment)
        self.assertEqual("Every paragraph that ends with a full stop like this one will be right aligned!\r" +
                         "This one will not!\r" +
                         "This one also will!", doc.get_text().strip())
        #ExEnd

    def test_delete_selection(self):

        #ExStart
        #ExFor:Node.range
        #ExFor:Range.delete
        #ExSummary:Shows how to delete all the nodes from a range.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Add text to the first section in the document, and then add another section.
        builder.write("Section 1. ")
        builder.insert_break(aw.BreakType.SECTION_BREAK_CONTINUOUS)
        builder.write("Section 2.")

        self.assertEqual("Section 1. \fSection 2.", doc.get_text().strip())

        # Remove the first section entirely by removing all the nodes
        # within its range, including the section itself.
        doc.sections[0].range.delete()

        self.assertEqual(1, doc.sections.count)
        self.assertEqual("Section 2.", doc.get_text().strip())
        #ExEnd

    def test_ranges_get_text(self):

        #ExStart
        #ExFor:Range
        #ExFor:Range.text
        #ExSummary:Shows how to get the text contents of all the nodes that a range covers.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Hello world!")

        self.assertEqual("Hello world!", doc.range.text.strip())
        #ExEnd

    ##ExStart
    ##ExFor:FindReplaceOptions.use_legacy_order
    ##ExSummary:Shows how to change the searching order of nodes when performing a find-and-replace text operation.
    #def test_use_legacy_order(self):

    #    for use_legacy_order in (True, False):
    #        with self.subTest(use_legacy_order=use_legacy_order):
    #            doc = aw.Document()
    #            builder = aw.DocumentBuilder(doc)

    #            # Insert three runs which we can search for using a regex pattern.
    #            # Place one of those runs inside a text box.
    #            builder.writeln("[tag 1]")
    #            text_box = builder.insert_shape(aw.drawing.ShapeType.TEXT_BOX, 100, 50)
    #            builder.writeln("[tag 2]")
    #            builder.move_to(text_box.first_paragraph)
    #            builder.write("[tag 3]")

    #            # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    #            options = aw.replacing.FindReplaceOptions()

    #            # Assign a custom callback to the "replacing_callback" property.
    #            callback = ExRange.TextReplacementTracker()
    #            options.replacing_callback = callback

    #            # If we set the "use_legacy_order" property to "True", the
    #            # find-and-replace operation will go through all the runs outside of a text box
    #            # before going through the ones inside a text box.
    #            # If we set the "use_legacy_order" property to "False", the
    #            # find-and-replace operation will go over all the runs in a range in sequential order.
    #            options.use_legacy_order = use_legacy_order

    #            doc.range.replace_regex(r"\[tag \d*\]", "", options)

    #            self.assertListEqual(
    #                ["[tag 1]", "[tag 3]", "[tag 2]"] if use_legacy_order else ["[tag 1]", "[tag 2]", "[tag 3]"],
    #                callback.matches)

    #class TextReplacementTracker(aw.replacing.IReplacingCallback):
    #    """Records the order of all matches that occur during a find-and-replace operation."""

    #    def __init__(self):
    #        self.matches: List[str] = []

    #    def replacing(self, e: aw.replacing.ReplacingArgs) -> aw.replacing.ReplaceAction:

    #        self.matches.add(e.match.value)
    #        return aw.replacing.ReplaceAction.REPLACE

    ##ExEnd

    def test_use_substitutions(self):

        for use_substitutions in (False, True):
            with self.subTest(use_substitutions=use_substitutions):
                #ExStart
                #ExFor:FindReplaceOptions.use_substitutions
                #ExSummary:Shows how to replace the text with substitutions.
                doc = aw.Document()
                builder = aw.DocumentBuilder(doc)

                builder.writeln("John sold a car to Paul.")
                builder.writeln("Jane sold a house to Joe.")

                # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
                options = aw.replacing.FindReplaceOptions()

                # Set the "use_substitutions" property to "True" to get
                # the find-and-replace operation to recognize substitution elements.
                # Set the "use_substitutions" property to "False" to ignore substitution elements.
                options.use_substitutions = use_substitutions

                regex = r"([A-z]+) sold a ([A-z]+) to ([A-z]+)"
                doc.range.replace_regex(regex, r"$3 bought a $2 from $1", options)

                if use_substitutions:
                    self.assertEqual(
                        "Paul bought a car from John.\rJoe bought a house from Jane.",
                        doc.get_text().strip())
                else:
                    self.assertEqual(
                        "$3 bought a $2 from $1.\r$3 bought a $2 from $1.",
                        doc.get_text().strip())
                #ExEnd

    ##ExStart
    ##ExFor:Range.replace(Regex,str,FindReplaceOptions)
    ##ExFor:IReplacingCallback
    ##ExFor:ReplaceAction
    ##ExFor:IReplacingCallback.replacing
    ##ExFor:ReplacingArgs
    ##ExFor:ReplacingArgs.match_node
    ##ExSummary:Shows how to insert an entire document's contents as a replacement of a match in a find-and-replace operation.
    #def test_insert_document_at_replace(self):

    #    main_doc = aw.Document(MY_DIR + "Document insertion destination.docx")

    #    # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    #    options = aw.replacing.FindReplaceOptions()
    #    options.replacing_callback = ExRange.InsertDocumentAtReplaceHandler()

    #    mainDoc.range.replace_regex(r"\[MY_DOCUMENT\]", "", options)
    #    mainDoc.save(ARTIFACTS_DIR + "InsertDocument.insert_document_at_replace.docx")

    #    self._test_insert_document_at_replace(aw.Document(ARTIFACTS_DIR + "InsertDocument.insert_document_at_replace.docx")); #ExSkip

    #class InsertDocumentAtReplaceHandler(aw.replacing.IReplacingCallback):

    #    def replacing(self, args: aw.replacing.ReplacingArgs) -> aw.replacing.ReplaceAction:

    #        sub_doc = aw.Document(MY_DIR + "Document.docx")

    #        # Insert a document after the paragraph containing the matched text.
    #        para = args.match_node.parent_node.as_paragraph()
    #        ExRange.insert_document(para, sub_doc)

    #        # Remove the paragraph with the matched text.
    #        para.remove()

    #        return aw.replacing.ReplaceAction.SKIP

    @staticmethod
    def insert_document(insertion_destination: aw.Node, doc_to_insert: aw.Document):
        """Inserts all the nodes of another document after a paragraph or table."""

        if insertion_destination.node_type == aw.NodeType.PARAGRAPH or insertion_destination.node_type == aw.NodeType.TABLE:

            dst_story = insertion_destination.parent_node

            importer = aw.NodeImporter(doc_to_insert, insertion_destination.document, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

            for src_section in doc_to_insert.sections:
                src_section = src_section.as_section()
                for src_node in src_section.body:

                    # Skip the node if it is the last empty paragraph in a section.
                    if src_node.node_type == aw.NodeType.PARAGRAPH:
                        para = src_node.as_paragraph()
                        if para.is_end_of_section and not para.has_child_nodes:
                            continue

                    new_node = importer.import_node(src_node, True)

                    dst_story.insert_after(new_node, insertion_destination)
                    insertion_destination = new_node
        else:
            raise ValueError("The destination node must be either a paragraph or table.")

    #ExEnd

    def _test_insert_document_at_replace(self, doc: aw.Document):

        self.assertEqual("1) At text that can be identified by regex:\rHello World!\r" +
                         "2) At a MERGEFIELD:\r\u0013 MERGEFIELD  Document_1  \\* MERGEFORMAT \u0014«Document_1»\u0015\r" +
                         "3) At a bookmark:", doc.first_section.body.get_text().strip())

    ##ExStart
    ##ExFor:FindReplaceOptions.direction
    ##ExFor:FindReplaceDirection
    ##ExSummary:Shows how to determine which direction a find-and-replace operation traverses the document in.
    #def test_direction(self):

    #    for find_replace_direction in (aw.replacing.FindReplaceDirection.BACKWARD, aw.replacing.FindReplaceDirection.FORWARD):
    #        with self.subTest(find_replace_direction=find_replace_direction):
    #            doc = aw.Document()
    #            builder = aw.DocumentBuilder(doc)

    #            # Insert three runs which we can search for using a regex pattern.
    #            # Place one of those runs inside a text box.
    #            builder.writeln("Match 1.")
    #            builder.writeln("Match 2.")
    #            builder.writeln("Match 3.")
    #            builder.writeln("Match 4.")

    #            # We can use a "FindReplaceOptions" object to modify the find-and-replace process.
    #            options = aw.replacing.FindReplaceOptions()

    #            # Assign a custom callback to the "replacing_callback" property.
    #            callback = ExRange.TextReplacementRecorder()
    #            options.replacing_callback = callback

    #            # Set the "direction" property to "FindReplaceDirection.BACKWARD" to get the find-and-replace
    #            # operation to start from the end of the range, and traverse back to the beginning.
    #            # Set the "direction" property to "FindReplaceDirection.FORWARD" to get the find-and-replace
    #            # operation to start from the beginning of the range, and traverse to the end.
    #            options.direction = find_replace_direction

    #            doc.range.replace_regex(r"Match \d*", "Replacement", options)

    #            self.assertEqual("Replacement.\r" +
    #                             "Replacement.\r" +
    #                             "Replacement.\r" +
    #                             "Replacement.", doc.get_text().strip())

    #            if find_replace_direction == aw.replacing.FindReplaceDirection.FORWARD:
    #                self.assertListEqual(["Match 1", "Match 2", "Match 3", "Match 4"], callback.matches)
    #            elif find_replace_direction == aw.replacing.FindReplaceDirection.BACKWARD:
    #                self.assertListEqual(["Match 4", "Match 3", "Match 2", "Match 1"], callback.matches)

    #class TextReplacementRecorder(aw.replacing.IReplacingCallback):
    #    """Records all matches that occur during a find-and-replace operation in the order that they take place."""

    #    def __init__(self):
    #        self.matches: List[str] = []

    #    def replacing(e: aw.replacing.ReplacingArgs) -> aw.replacing.ReplaceAction:

    #        self.matches.add(e.match.value)
    #        return aw.replacing.ReplaceAction.REPLACE

    ##ExEnd
