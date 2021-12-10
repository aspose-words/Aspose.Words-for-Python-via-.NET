import unittest
import io
import os
from datetime import datetime, timedelta, timezone

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir, image_dir
from document_helper import DocumentHelper
from testutil import TestUtil

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir
IMAGE_DIR = image_dir

class ExDocumentProperties(ApiExampleBase):

    def test_built_in(self):

        #ExStart
        #ExFor:BuiltInDocumentProperties
        #ExFor:Document.BuiltInDocumentProperties
        #ExFor:Document.CustomDocumentProperties
        #ExFor:DocumentProperty
        #ExFor:DocumentProperty.Name
        #ExFor:DocumentProperty.Value
        #ExFor:DocumentProperty.Type
        #ExSummary:Shows how to work with built-in document properties.
        doc = aw.Document(MY_DIR + "Properties.docx")

        # The "Document" object contains some of its metadata in its members.
        print(f"Document filename:\n\t \"{doc.original_file_name}\"")

        # The document also stores metadata in its built-in properties.
        # Each built-in property is a member of the document's "BuiltInDocumentProperties" object.
        print("Built-in Properties:")
        for doc_property in doc.built_in_document_properties:
            print(doc_property.name)
            print(f"\tType:\t{doc_property.type}")

            # Some properties may store multiple values.
            if isinstance(doc_property.value, list):
                for value in doc_property.value:
                    print(f"\tValue:\t\"{value}\"")
            else:
                print(f"\tValue:\t\"{doc_property.value}\"")

        #ExEnd

        self.assertEqual(28, doc.built_in_document_properties.count)

    def test_custom(self):

        #ExStart
        #ExFor:BuiltInDocumentProperties.Item(String)
        #ExFor:CustomDocumentProperties
        #ExFor:DocumentProperty.ToString
        #ExFor:DocumentPropertyCollection.Count
        #ExFor:DocumentPropertyCollection.Item(int)
        #ExSummary:Shows how to work with custom document properties.
        doc = aw.Document(MY_DIR + "Properties.docx")

        # Every document contains a collection of custom properties, which, like the built-in properties, are key-value pairs.
        # The document has a fixed list of built-in properties. The user creates all of the custom properties.
        self.assertEqual("Value of custom document property", str(doc.custom_document_properties.get_by_name("CustomProperty")))

        doc.custom_document_properties.add("CustomProperty2", "Value of custom document property #2")

        print("Custom Properties:")
        for custom_document_property in doc.custom_document_properties:
            print(custom_document_property.name)
            print(f"\tType:\t{custom_document_property.type}")
            print(f"\tValue:\t\"{custom_document_property.value}\"")

        #ExEnd

        self.assertEqual(2, doc.custom_document_properties.count)

    def test_description(self):

        #ExStart
        #ExFor:BuiltInDocumentProperties.Author
        #ExFor:BuiltInDocumentProperties.Category
        #ExFor:BuiltInDocumentProperties.Comments
        #ExFor:BuiltInDocumentProperties.Keywords
        #ExFor:BuiltInDocumentProperties.Subject
        #ExFor:BuiltInDocumentProperties.Title
        #ExSummary:Shows how to work with built-in document properties in the "Description" category.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        properties = doc.built_in_document_properties

        # Below are four built-in document properties that have fields that can display their values in the document body.
        # 1 -  "author" property, which we can display using an AUTHOR field:
        properties.author = "John Doe"
        builder.write("Author:\t")
        builder.insert_field(aw.fields.FieldType.FIELD_AUTHOR, True)

        # 2 -  "title" property, which we can display using a TITLE field:
        properties.title = "John's Document"
        builder.write("\nDoc title:\t")
        builder.insert_field(aw.fields.FieldType.FIELD_TITLE, True)

        # 3 -  "subject" property, which we can display using a SUBJECT field:
        properties.subject = "My subject"
        builder.write("\nSubject:\t")
        builder.insert_field(aw.fields.FieldType.FIELD_SUBJECT, True)

        # 4 -  "comments" property, which we can display using a COMMENTS field:
        properties.comments = f"This is {properties.author}'s document about {properties.subject}"
        builder.write("\nComments:\t\"")
        builder.insert_field(aw.fields.FieldType.FIELD_COMMENTS, True)
        builder.write("\"")

        # The "category" built-in property does not have a field that can display its value.
        properties.category = "My category"

        # We can set multiple keywords for a document by separating the string value of the "keywords" property with semicolons.
        properties.keywords = "Tag 1; Tag 2; Tag 3"

        # We can right-click this document in Windows Explorer and find these properties in "Properties" -> "Details".
        # The "Author" built-in property is in the "Origin" group, and the others are in the "Description" group.
        doc.save(ARTIFACTS_DIR + "DocumentProperties.description.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "DocumentProperties.description.docx")

        properties = doc.built_in_document_properties

        self.assertEqual("John Doe", properties.author)
        self.assertEqual("My category", properties.category)
        self.assertEqual(f"This is {properties.author}'s document about {properties.subject}", properties.comments)
        self.assertEqual("Tag 1; Tag 2; Tag 3", properties.keywords)
        self.assertEqual("My subject", properties.subject)
        self.assertEqual("John's Document", properties.title)
        self.assertEqual("Author:\t\u0013 AUTHOR \u0014John Doe\u0015\r" +
                         "Doc title:\t\u0013 TITLE \u0014John's Document\u0015\r" +
                         "Subject:\t\u0013 SUBJECT \u0014My subject\u0015\r" +
                         "Comments:\t\"\u0013 COMMENTS \u0014This is John Doe's document about My subject\u0015\"", doc.get_text().strip())

    def test_origin(self):

        #ExStart
        #ExFor:BuiltInDocumentProperties.Company
        #ExFor:BuiltInDocumentProperties.CreatedTime
        #ExFor:BuiltInDocumentProperties.LastPrinted
        #ExFor:BuiltInDocumentProperties.LastSavedBy
        #ExFor:BuiltInDocumentProperties.LastSavedTime
        #ExFor:BuiltInDocumentProperties.Manager
        #ExFor:BuiltInDocumentProperties.NameOfApplication
        #ExFor:BuiltInDocumentProperties.RevisionNumber
        #ExFor:BuiltInDocumentProperties.Template
        #ExFor:BuiltInDocumentProperties.TotalEditingTime
        #ExFor:BuiltInDocumentProperties.Version
        #ExSummary:Shows how to work with document properties in the "Origin" category.
        # Open a document that we have created and edited using Microsoft Word.
        doc = aw.Document(MY_DIR + "Properties.docx")
        properties = doc.built_in_document_properties

        # The following built-in properties contain information regarding the creation and editing of this document.
        # We can right-click this document in Windows Explorer and find
        # these properties via "Properties" -> "Details" -> "Origin" category.
        # Fields such as PRINTDATE and EDITTIME can display these values in the document body.
        print(f"Created using {properties.name_of_application}, on {properties.created_time}")
        print(f"Minutes spent editing: {properties.total_editing_time}")
        print(f"Date/time last printed: {properties.last_printed}")
        print(f"Template document: {properties.template}")

        # We can also change the values of built-in properties.
        properties.company = "Doe Ltd."
        properties.manager = "Jane Doe"
        properties.version = 5
        properties.revision_number += 1

        # Microsoft Word updates the following properties automatically when we save the document.
        # To use these properties with Aspose.Words, we will need to set values for them manually.
        properties.last_saved_by = "John Doe"
        properties.last_saved_time = datetime.utcnow()

        # We can right-click this document in Windows Explorer and find these properties in "Properties" -> "Details" -> "Origin".
        doc.save(ARTIFACTS_DIR + "DocumentProperties.origin.docx")
        #ExEnd

        properties = aw.Document(ARTIFACTS_DIR + "DocumentProperties.origin.docx").built_in_document_properties

        self.assertEqual("Doe Ltd.", properties.company)
        self.assertEqual(datetime(2006, 4, 25, 10, 10, 0, tzinfo=timezone.utc), properties.created_time)
        self.assertEqual(datetime(2019, 4, 21, 10, 0, 0, tzinfo=timezone.utc), properties.last_printed)
        self.assertEqual("John Doe", properties.last_saved_by)
        TestUtil.verify_date(self, datetime.now(timezone.utc), properties.last_saved_time, timedelta(seconds=5))
        self.assertEqual("Jane Doe", properties.manager)
        self.assertEqual("Microsoft Office Word", properties.name_of_application)
        self.assertEqual(12, properties.revision_number)
        self.assertEqual("Normal", properties.template)
        self.assertEqual(8, properties.total_editing_time)
        self.assertEqual(786432, properties.version)

    #ExStart
    #ExFor:BuiltInDocumentProperties.Bytes
    #ExFor:BuiltInDocumentProperties.Characters
    #ExFor:BuiltInDocumentProperties.CharactersWithSpaces
    #ExFor:BuiltInDocumentProperties.ContentStatus
    #ExFor:BuiltInDocumentProperties.ContentType
    #ExFor:BuiltInDocumentProperties.Lines
    #ExFor:BuiltInDocumentProperties.LinksUpToDate
    #ExFor:BuiltInDocumentProperties.Pages
    #ExFor:BuiltInDocumentProperties.Paragraphs
    #ExFor:BuiltInDocumentProperties.Words
    #ExSummary:Shows how to work with document properties in the "Content" category.
    def test_content(self):

        doc = aw.Document(MY_DIR + "Paragraphs.docx")
        properties = doc.built_in_document_properties

        # By using built in properties,
        # we can treat document statistics such as word/page/character counts as metadata that can be glanced at without opening the document
        # These properties are accessed by right clicking the file in Windows Explorer and navigating to Properties > Details > Content
        # If we want to display this data inside the document, we can use fields such as NUMPAGES, NUMWORDS, NUMCHARS etc.
        # Also, these values can also be viewed in Microsoft Word by navigating File > Properties > Advanced Properties > Statistics
        # Page count: The page_count property shows the page count in real time and its value can be assigned to the Pages property

        # The "pages" property stores the page count of the document.
        self.assertEqual(6, properties.pages)

        # The "words", "characters", and "characters_with_spaces" built-in properties also display various document statistics,
        # but we need to call the "update_word_count" method on the whole document before we can expect them to contain accurate values.
        self.assertEqual(1054, properties.words) #ExSkip
        self.assertEqual(6009, properties.characters) #ExSkip
        self.assertEqual(7049, properties.characters_with_spaces) #ExSkip
        doc.update_word_count()

        self.assertEqual(1035, properties.words)
        self.assertEqual(6026, properties.characters)
        self.assertEqual(7041, properties.characters_with_spaces)

        # Count the number of lines in the document, and then assign the result to the "Lines" built-in property.
        line_counter = ExDocumentProperties.LineCounter(doc)
        properties.lines = line_counter.get_line_count()

        self.assertEqual(142, properties.lines)

        # Assign the number of Paragraph nodes in the document to the "paragraphs" built-in property.
        properties.paragraphs = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True).count
        self.assertEqual(29, properties.paragraphs)

        # Get an estimate of the file size of our document via the "bytes" built-in property.
        self.assertEqual(20310, properties.bytes)

        # Set a different template for our document, and then update the "template" built-in property manually to reflect this change.
        doc.attached_template = MY_DIR + "Business brochure.dotx"

        self.assertEqual("Normal", properties.template)

        properties.template = doc.attached_template

        # "content_status" is a descriptive built-in property.
        properties.content_status = "Draft"

        # Upon saving, the "content_type" built-in property will contain the MIME type of the output save format.
        self.assertEqual("", properties.content_type)

        # If the document contains links, and they are all up to date, we can set the "links_up_to_date" property to "True".
        self.assertFalse(properties.links_up_to_date)

        doc.save(ARTIFACTS_DIR + "DocumentProperties.content.docx")
        self._test_content(aw.Document(ARTIFACTS_DIR + "DocumentProperties.content.docx")) #ExSkip

    class LineCounter:
        """Counts the lines in a document.
        Traverses the document's layout entities tree upon construction,
        counting entities of the "Line" type that also contain real text."""

        def __init__(self, doc: aw.Document):

            self.layout_enumerator = aw.layout.LayoutEnumerator(doc)
            self.line_count = 0
            self.scanning_line_for_real_text = False

            self.count_lines()

        def get_line_count(self) -> int:

            return self.line_count

        def count_lines(self) -> int:

            while True:
                if self.layout_enumerator.type == aw.layout.LayoutEntityType.LINE:
                    self.scanning_line_for_real_text = True

                if self.layout_enumerator.move_first_child():
                    if self.scanning_line_for_real_text and self.layout_enumerator.kind.startswith("TEXT"):
                        self.line_count += 1
                        self.scanning_line_for_real_text = False

                    self.count_lines()
                    self.layout_enumerator.move_parent()

                if not self.layout_enumerator.move_next():
                    break

    #ExEnd

    def _test_content(self, doc: aw.Document):

        properties = doc.built_in_document_properties

        self.assertEqual(6, properties.pages)

        self.assertEqual(1035, properties.words)
        self.assertEqual(6026, properties.characters)
        self.assertEqual(7041, properties.characters_with_spaces)
        self.assertEqual(142, properties.lines)
        self.assertEqual(29, properties.paragraphs)
        self.assertAlmostEqual(15500, properties.bytes, delta=200)
        self.assertEqual(MY_DIR.replace("\\\\", "\\") + "Business brochure.dotx", properties.template)
        self.assertEqual("Draft", properties.content_status)
        self.assertEqual("", properties.content_type)
        self.assertFalse(properties.links_up_to_date)

    def test_thumbnail(self):

        #ExStart
        #ExFor:BuiltInDocumentProperties.Thumbnail
        #ExFor:DocumentProperty.ToByteArray
        #ExSummary:Shows how to add a thumbnail to a document that we save as an Epub.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
        builder.writeln("Hello world!")

        # If we save a document, whose "Thumbnail" property contains image data that we added, as an Epub,
        # a reader that opens that document may display the image before the first page.
        properties = doc.built_in_document_properties

        with open(IMAGE_DIR + "Logo.jpg", "rb") as file:
            thumbnail_bytes = file.read()

        properties.thumbnail = thumbnail_bytes

        doc.save(ARTIFACTS_DIR + "DocumentProperties.thumbnail.epub")

        # We can extract a document's thumbnail image and save it to the local file system.
        thumbnail = doc.built_in_document_properties.get_by_name("Thumbnail")

        with open(ARTIFACTS_DIR + "DocumentProperties.thumbnail.gif", "wb") as file:
            file.write(thumbnail.to_byte_array())
        #ExEnd

        with open(ARTIFACTS_DIR + "DocumentProperties.thumbnail.gif", "rb") as img_stream:
            TestUtil.verify_image(self, 400, 400, image_stream=img_stream)

    def test_hyperlink_base(self):

        #ExStart
        #ExFor:BuiltInDocumentProperties.HyperlinkBase
        #ExSummary:Shows how to store the base part of a hyperlink in the document's properties.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert a relative hyperlink to a document in the local file system named "Document.docx".
        # Clicking on the link in Microsoft Word will open the designated document, if it is available.
        builder.insert_hyperlink("Relative hyperlink", "Document.docx", False)

        # This link is relative. If there is no "Document.docx" in the same folder
        # as the document that contains this link, the link will be broken.
        self.assertFalse(os.path.exists(ARTIFACTS_DIR + "Document.docx"))
        doc.save(ARTIFACTS_DIR + "DocumentProperties.hyperlink_base.broken_link.docx")

        # The document we are trying to link to is in a different directory to the one we are planning to save the document in.
        # We could fix links like this by putting an absolute filename in each one.
        # Alternatively, we could provide a base link that every hyperlink with a relative filename
        # will prepend to its link when we click on it.
        properties = doc.built_in_document_properties
        properties.hyperlink_base = MY_DIR

        self.assertTrue(os.path.exists(properties.hyperlink_base + doc.range.fields[0].as_field_hyperlink().address))

        doc.save(ARTIFACTS_DIR + "DocumentProperties.hyperlink_base.working_link.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "DocumentProperties.hyperlink_base.broken_link.docx")
        properties = doc.built_in_document_properties

        self.assertEqual("", properties.hyperlink_base)

        doc = aw.Document(ARTIFACTS_DIR + "DocumentProperties.hyperlink_base.working_link.docx")
        properties = doc.built_in_document_properties

        self.assertEqual(MY_DIR, properties.hyperlink_base)
        self.assertTrue(os.path.exists(properties.hyperlink_base + doc.range.fields[0].as_field_hyperlink().address))

    def test_heading_pairs(self):

        #ExStart
        #ExFor:Properties.BuiltInDocumentProperties.HeadingPairs
        #ExFor:Properties.BuiltInDocumentProperties.TitlesOfParts
        #ExSummary:Shows the relationship between "HeadingPairs" and "TitlesOfParts" properties.
        doc = aw.Document(MY_DIR + "Heading pairs and titles of parts.docx")

        # We can find the combined values of these collections via
        # "File" -> "Properties" -> "Advanced Properties" -> "Contents" tab.
        # The HeadingPairs property is a collection of <string, int> pairs that
        # determines how many document parts a heading spans across.
        heading_pairs = doc.built_in_document_properties.heading_pairs

        # The titles_of_parts property contains the names of parts that belong to the above headings.
        titles_of_parts = doc.built_in_document_properties.titles_of_parts

        heading_pairs_index = 0
        titles_of_parts_index = 0
        while heading_pairs_index < len(heading_pairs):

            print(f"Parts for {heading_pairs[heading_pairs_index]}:")
            heading_pairs_index += 1
            parts_count = int(heading_pairs[heading_pairs_index])
            heading_pairs_index += 1

            for i in range(parts_count):
                print(f"\t\"{titles_of_parts[titles_of_parts_index]}\"")
                titles_of_parts_index += 1

        #ExEnd

        # There are 6 array elements designating 3 heading/part count pairs
        self.assertEqual(6, len(heading_pairs))
        self.assertEqual("Title", heading_pairs[0])
        self.assertEqual(1, heading_pairs[1])
        self.assertEqual("Heading 1", heading_pairs[2])
        self.assertEqual(5, heading_pairs[3])
        self.assertEqual("Heading 2", heading_pairs[4])
        self.assertEqual(2, heading_pairs[5])

        self.assertEqual(8, len(titles_of_parts))
        # "Title"
        self.assertEqual("", titles_of_parts[0])
        # "Heading 1"
        self.assertEqual("Part1", titles_of_parts[1])
        self.assertEqual("Part2", titles_of_parts[2])
        self.assertEqual("Part3", titles_of_parts[3])
        self.assertEqual("Part4", titles_of_parts[4])
        self.assertEqual("Part5", titles_of_parts[5])
        # "Heading 2"
        self.assertEqual("Part6", titles_of_parts[6])
        self.assertEqual("Part7", titles_of_parts[7])

    def test_security(self):

        #ExStart
        #ExFor:Properties.BuiltInDocumentProperties.Security
        #ExFor:Properties.DocumentSecurity
        #ExSummary:Shows how to use document properties to display the security level of a document.
        doc = aw.Document()

        self.assertEqual(aw.properties.DocumentSecurity.NONE, doc.built_in_document_properties.security)

        # If we configure a document to be read-only, it will display this status using the "security" built-in property.
        doc.write_protection.read_only_recommended = True
        doc.save(ARTIFACTS_DIR + "DocumentProperties.security.read_only_recommended.docx")

        self.assertEqual(aw.properties.DocumentSecurity.READ_ONLY_RECOMMENDED,
            aw.Document(ARTIFACTS_DIR + "DocumentProperties.security.read_only_recommended.docx").built_in_document_properties.security)

        # Write-protect a document, and then verify its security level.
        doc = aw.Document()

        self.assertFalse(doc.write_protection.is_write_protected)

        doc.write_protection.set_password("MyPassword")

        self.assertTrue(doc.write_protection.validate_password("MyPassword"))
        self.assertTrue(doc.write_protection.is_write_protected)

        doc.save(ARTIFACTS_DIR + "DocumentProperties.security.read_only_enforced.docx")

        self.assertEqual(aw.properties.DocumentSecurity.READ_ONLY_ENFORCED,
            aw.Document(ARTIFACTS_DIR + "DocumentProperties.security.read_only_enforced.docx").built_in_document_properties.security)

        # "security" is a descriptive property. We can edit its value manually.
        doc = aw.Document()

        doc.protect(aw.ProtectionType.ALLOW_ONLY_COMMENTS, "MyPassword")
        doc.built_in_document_properties.security = aw.properties.DocumentSecurity.READ_ONLY_EXCEPT_ANNOTATIONS
        doc.save(ARTIFACTS_DIR + "DocumentProperties.security.read_only_except_annotations.docx")

        self.assertEqual(aw.properties.DocumentSecurity.READ_ONLY_EXCEPT_ANNOTATIONS,
            aw.Document(ARTIFACTS_DIR + "DocumentProperties.security.read_only_except_annotations.docx").built_in_document_properties.security)
        #ExEnd

    def test_custom_named_access(self):

        #ExStart
        #ExFor:DocumentPropertyCollection.Item(String)
        #ExFor:CustomDocumentProperties.Add(String,DateTime)
        #ExFor:DocumentProperty.ToDateTime
        #ExSummary:Shows how to create a custom document property which contains a date and time.
        doc = aw.Document()

        doc.custom_document_properties.add("AuthorizationDate", datetime.utcnow())

        print("Document authorized on", doc.custom_document_properties.get_by_name("AuthorizationDate").to_date_time())
        #ExEnd

        TestUtil.verify_date(self, datetime.now(tz=timezone.utc),
            DocumentHelper.save_open(doc).custom_document_properties.get_by_name("AuthorizationDate").to_date_time(),
            timedelta(seconds=1))

    def test_link_custom_document_properties_to_bookmark(self):

        #ExStart
        #ExFor:CustomDocumentProperties.AddLinkToContent(String, String)
        #ExFor:DocumentProperty.IsLinkToContent
        #ExFor:DocumentProperty.LinkSource
        #ExSummary:Shows how to link a custom document property to a bookmark.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.start_bookmark("MyBookmark")
        builder.write("Hello world!")
        builder.end_bookmark("MyBookmark")

        # Link a new custom property to a bookmark. The value of this property
        # will be the contents of the bookmark that it references in the "link_source" member.
        custom_properties = doc.custom_document_properties
        custom_property = custom_properties.add_link_to_content("Bookmark", "MyBookmark")

        self.assertEqual(True, custom_property.is_link_to_content)
        self.assertEqual("MyBookmark", custom_property.link_source)
        self.assertEqual("Hello world!", custom_property.value)

        doc.save(ARTIFACTS_DIR + "DocumentProperties.link_custom_document_properties_to_bookmark.docx")
        #ExEnd

        doc = aw.Document(ARTIFACTS_DIR + "DocumentProperties.link_custom_document_properties_to_bookmark.docx")
        custom_property = doc.custom_document_properties.get_by_name("Bookmark")

        self.assertEqual(True, custom_property.is_link_to_content)
        self.assertEqual("MyBookmark", custom_property.link_source)
        self.assertEqual("Hello world!", custom_property.value)

    def test_document_property_collection(self):

        #ExStart
        #ExFor:CustomDocumentProperties.Add(String,String)
        #ExFor:CustomDocumentProperties.Add(String,Boolean)
        #ExFor:CustomDocumentProperties.Add(String,int)
        #ExFor:CustomDocumentProperties.Add(String,DateTime)
        #ExFor:CustomDocumentProperties.Add(String,Double)
        #ExFor:DocumentProperty.Type
        #ExFor:Properties.DocumentPropertyCollection
        #ExFor:Properties.DocumentPropertyCollection.Clear
        #ExFor:Properties.DocumentPropertyCollection.Contains(System.String)
        #ExFor:Properties.DocumentPropertyCollection.GetEnumerator
        #ExFor:Properties.DocumentPropertyCollection.IndexOf(System.String)
        #ExFor:Properties.DocumentPropertyCollection.RemoveAt(System.Int32)
        #ExFor:Properties.DocumentPropertyCollection.Remove
        #ExFor:PropertyType
        #ExSummary:Shows how to work with a document's custom properties.
        doc = aw.Document()
        properties = doc.custom_document_properties

        self.assertEqual(0, properties.count)

        # Custom document properties are key-value pairs that we can add to the document.
        properties.add("Authorized", True)
        properties.add("Authorized By", "John Doe")
        properties.add("Authorized Date", datetime.now())
        properties.add("Authorized Revision", doc.built_in_document_properties.revision_number)
        properties.add("Authorized Amount", 123.45)

        # The collection sorts the custom properties in alphabetic order.
        self.assertEqual(1, properties.index_of("Authorized Amount"))
        self.assertEqual(5, properties.count)

        # Print every custom property in the document.
        for prop in properties:
            print(f"Name: \"{prop.name}\"\n\tType: \"{prop.type}\"\n\tValue: \"{prop.value}\"")

        # Display the value of a custom property using a DOCPROPERTY field.
        builder = aw.DocumentBuilder(doc)
        field = builder.insert_field(" DOCPROPERTY \"Authorized By\"").as_field_doc_property()
        field.update()

        self.assertEqual("John Doe", field.result)

        # We can find these custom properties in Microsoft Word via "File" -> "Properties" > "Advanced Properties" > "Custom".
        doc.save(ARTIFACTS_DIR + "DocumentProperties.document_property_collection.docx")

        # Below are three ways or removing custom properties from a document.
        # 1 -  Remove by index:
        properties.remove_at(1)

        self.assertFalse(properties.contains("Authorized Amount"))
        self.assertEqual(4, properties.count)

        # 2 -  Remove by name:
        properties.remove("Authorized Revision")

        self.assertFalse(properties.contains("Authorized Revision"))
        self.assertEqual(3, properties.count)

        # 3 -  Empty the entire collection at once:
        properties.clear()

        self.assertEqual(0, properties.count)
        #ExEnd

    def test_property_types(self):

        #ExStart
        #ExFor:DocumentProperty.ToBool
        #ExFor:DocumentProperty.ToInt
        #ExFor:DocumentProperty.ToDouble
        #ExFor:DocumentProperty.ToString
        #ExFor:DocumentProperty.ToDateTime
        #ExSummary:Shows various type conversion methods of custom document properties.
        doc = aw.Document()
        properties = doc.custom_document_properties

        auth_date = datetime.now()
        properties.add("Authorized", True)
        properties.add("Authorized By", "John Doe")
        properties.add("Authorized Date", auth_date)
        properties.add("Authorized Revision", doc.built_in_document_properties.revision_number)
        properties.add("Authorized Amount", 123.45)

        self.assertEqual(True, properties.get_by_name("Authorized").to_bool())
        self.assertEqual("John Doe", str(properties.get_by_name("Authorized By")))
        self.assertEqual(auth_date, properties.get_by_name("Authorized Date").to_date_time())
        self.assertEqual(1, properties.get_by_name("Authorized Revision").to_int())
        self.assertEqual(123.45, properties.get_by_name("Authorized Amount").to_double())
        #ExEnd
