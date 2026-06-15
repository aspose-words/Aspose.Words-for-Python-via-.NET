import os
from document_helper import DocumentHelper
from document_helper import DocumentHelper
from datetime import timedelta, timezone
import sys
# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
import aspose.words.fields
import aspose.words.layout
import aspose.words.properties
import datetime
import system_helper
import test_util
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, IMAGE_DIR, MY_DIR

class ExDocumentProperties(ApiExampleBase):

    def test_built_in(self):
        #ExStart
        #ExFor:BuiltInDocumentProperties
        #ExFor:Document.built_in_document_properties
        #ExFor:Document.custom_document_properties
        #ExFor:DocumentProperty
        #ExFor:DocumentProperty.name
        #ExFor:DocumentProperty.value
        #ExFor:DocumentProperty.type
        #ExSummary:Shows how to work with built-in document properties.
        doc = aw.Document(file_name=MY_DIR + 'Properties.docx')
        # The "Document" object contains some of its metadata in its members.
        print(f'Document filename:\n\t "{doc.original_file_name}"')
        # The document also stores metadata in its built-in properties.
        # Each built-in property is a member of the document's "BuiltInDocumentProperties" object.
        print('Built-in Properties:')
        for doc_property in doc.built_in_document_properties:
            print(doc_property.name)
            print(f'\tType:\t{doc_property.type}')
            # Some properties may store multiple values.
            if isinstance(doc_property.value, (list, tuple)):
                for value in doc_property.value:
                    print(f'\tValue:\t"{value}"')
            else:
                print(f'\tValue:\t"{doc_property.value}"')
        #ExEnd
        self.assertEqual(31, doc.built_in_document_properties.count)

    def test_description(self):
        #ExStart
        #ExFor:BuiltInDocumentProperties.author
        #ExFor:BuiltInDocumentProperties.category
        #ExFor:BuiltInDocumentProperties.comments
        #ExFor:BuiltInDocumentProperties.keywords
        #ExFor:BuiltInDocumentProperties.subject
        #ExFor:BuiltInDocumentProperties.title
        #ExSummary:Shows how to work with built-in document properties in the "Description" category.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        properties = doc.built_in_document_properties
        # Below are four built-in document properties that have fields that can display their values in the document body.
        # 1 -  "Author" property, which we can display using an AUTHOR field:
        properties.author = 'John Doe'
        builder.write('Author:\t')
        builder.insert_field(field_type=aw.fields.FieldType.FIELD_AUTHOR, update_field=True)
        # 2 -  "Title" property, which we can display using a TITLE field:
        properties.title = "John's Document"
        builder.write('\nDoc title:\t')
        builder.insert_field(field_type=aw.fields.FieldType.FIELD_TITLE, update_field=True)
        # 3 -  "Subject" property, which we can display using a SUBJECT field:
        properties.subject = 'My subject'
        builder.write('\nSubject:\t')
        builder.insert_field(field_type=aw.fields.FieldType.FIELD_SUBJECT, update_field=True)
        # 4 -  "Comments" property, which we can display using a COMMENTS field:
        properties.comments = f"This is {properties.author}'s document about {properties.subject}"
        builder.write('\nComments:\t"')
        builder.insert_field(field_type=aw.fields.FieldType.FIELD_COMMENTS, update_field=True)
        builder.write('"')
        # The "Category" built-in property does not have a field that can display its value.
        properties.category = 'My category'
        # We can set multiple keywords for a document by separating the string value of the "Keywords" property with semicolons.
        properties.keywords = 'Tag 1; Tag 2; Tag 3'
        # We can right-click this document in Windows Explorer and find these properties in "Properties" -> "Details".
        # The "Author" built-in property is in the "Origin" group, and the others are in the "Description" group.
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentProperties.Description.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentProperties.Description.docx')
        properties = doc.built_in_document_properties
        self.assertEqual('John Doe', properties.author)
        self.assertEqual('My category', properties.category)
        self.assertEqual(f"This is {properties.author}'s document about {properties.subject}", properties.comments)
        self.assertEqual('Tag 1; Tag 2; Tag 3', properties.keywords)
        self.assertEqual('My subject', properties.subject)
        self.assertEqual("John's Document", properties.title)
        self.assertEqual('Author:\t\x13 AUTHOR \x14John Doe\x15\r' + "Doc title:\t\x13 TITLE \x14John's Document\x15\r" + 'Subject:\t\x13 SUBJECT \x14My subject\x15\r' + 'Comments:\t"\x13 COMMENTS \x14This is John Doe\'s document about My subject\x15"', doc.get_text().strip())

    def _test_content(self, doc):
        properties = doc.built_in_document_properties
        self.assertEqual(6, properties.pages)
        self.assertEqual(1035, properties.words)
        self.assertEqual(6026, properties.characters)
        self.assertEqual(7041, properties.characters_with_spaces)
        self.assertEqual(142, properties.lines)
        self.assertEqual(29, properties.paragraphs)
        self.assertAlmostEqual(15500, properties.bytes, delta=200)
        self.assertEqual(MY_DIR.replace('\\\\', '\\') + 'Business brochure.dotx', properties.template)
        self.assertEqual('Draft', properties.content_status)
        self.assertEqual('', properties.content_type)
        self.assertFalse(properties.links_up_to_date)

    def test_thumbnail(self):
        #ExStart
        #ExFor:BuiltInDocumentProperties.thumbnail
        #ExFor:DocumentProperty.to_byte_array
        #ExSummary:Shows how to add a thumbnail to a document that we save as an Epub.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('Hello world!')
        # If we save a document, whose "Thumbnail" property contains image data that we added, as an Epub,
        # a reader that opens that document may display the image before the first page.
        properties = doc.built_in_document_properties
        thumbnail_bytes = system_helper.io.File.read_all_bytes(IMAGE_DIR + 'Logo.jpg')
        properties.thumbnail = thumbnail_bytes
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentProperties.Thumbnail.epub')
        # We can extract a document's thumbnail image and save it to the local file system.
        thumbnail = doc.built_in_document_properties.get_by_name('Thumbnail')
        system_helper.io.File.write_all_bytes(ARTIFACTS_DIR + 'DocumentProperties.Thumbnail.gif', thumbnail.to_byte_array())
        #ExEnd
        test_util.TestUtil.verify_image(400, 400, ARTIFACTS_DIR + 'DocumentProperties.Thumbnail.gif')

    @unittest.skipIf(sys.platform.startswith('win'), 'Discrepancy in assertion between Python and .Net')
    def test_hyperlink_base(self):
        #ExStart
        #ExFor:BuiltInDocumentProperties.hyperlink_base
        #ExSummary:Shows how to store the base part of a hyperlink in the document's properties.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        # Insert a relative hyperlink to a document in the local file system named "Document.docx".
        # Clicking on the link in Microsoft Word will open the designated document, if it is available.
        builder.insert_hyperlink('Relative hyperlink', 'Document.docx', False)
        # This link is relative. If there is no "Document.docx" in the same folder
        # as the document that contains this link, the link will be broken.
        self.assertFalse(system_helper.io.File.exist(ARTIFACTS_DIR + 'Document.docx'))
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentProperties.HyperlinkBase.BrokenLink.docx')
        # The document we are trying to link to is in a different directory to the one we are planning to save the document in.
        # We could fix links like this by putting an absolute filename in each one.
        # Alternatively, we could provide a base link that every hyperlink with a relative filename
        # will prepend to its link when we click on it.
        properties = doc.built_in_document_properties
        properties.hyperlink_base = MY_DIR
        self.assertTrue(system_helper.io.File.exist(properties.hyperlink_base + doc.range.fields[0].as_field_hyperlink().address))
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentProperties.HyperlinkBase.WorkingLink.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentProperties.HyperlinkBase.BrokenLink.docx')
        properties = doc.built_in_document_properties
        self.assertEqual('', properties.hyperlink_base)
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentProperties.HyperlinkBase.WorkingLink.docx')
        properties = doc.built_in_document_properties
        self.assertEqual(MY_DIR, properties.hyperlink_base)
        self.assertTrue(system_helper.io.File.exist(properties.hyperlink_base + doc.range.fields[0].as_field_hyperlink().address))

    def test_heading_pairs(self):
        from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, GOLDS_DIR, TEMP_DIR, IMAGE_DIR, FONTS_DIR
        import aspose.words as aw
        doc = aw.Document(file_name=MY_DIR + 'Heading pairs and titles of parts.docx')
        heading_pairs = doc.built_in_document_properties.heading_pairs
        titles_of_parts = doc.built_in_document_properties.titles_of_parts
        heading_pairs_index = 0
        titles_of_parts_index = 0
        while heading_pairs_index < len(heading_pairs):
            print(f'Parts for {heading_pairs[heading_pairs_index]}:')
            heading_pairs_index += 1
            parts_count = int(heading_pairs[heading_pairs_index])
            heading_pairs_index += 1
            for i in range(parts_count):
                print(f'\t"{titles_of_parts[titles_of_parts_index]}"')
                titles_of_parts_index += 1
        # There are 6 array elements designating 3 heading/part count pairs
        self.assertEqual(6, len(heading_pairs))
        self.assertEqual('Title', str(heading_pairs[0]))
        self.assertEqual('1', str(heading_pairs[1]))
        self.assertEqual('Heading 1', str(heading_pairs[2]))
        self.assertEqual('5', str(heading_pairs[3]))
        self.assertEqual('Heading 2', str(heading_pairs[4]))
        self.assertEqual('2', str(heading_pairs[5]))
        self.assertEqual(8, len(titles_of_parts))
        # "Title"
        self.assertEqual('', titles_of_parts[0])
        # "Heading 1"
        self.assertEqual('Part1', titles_of_parts[1])
        self.assertEqual('Part2', titles_of_parts[2])
        self.assertEqual('Part3', titles_of_parts[3])
        self.assertEqual('Part4', titles_of_parts[4])
        self.assertEqual('Part5', titles_of_parts[5])
        # "Heading 2"
        self.assertEqual('Part6', titles_of_parts[6])
        self.assertEqual('Part7', titles_of_parts[7])

    def test_security(self):
        #ExStart
        #ExFor:BuiltInDocumentProperties.security
        #ExFor:DocumentSecurity
        #ExSummary:Shows how to use document properties to display the security level of a document.
        doc = aw.Document()
        self.assertEqual(aw.properties.DocumentSecurity.NONE, doc.built_in_document_properties.security)
        # If we configure a document to be read-only, it will display this status using the "Security" built-in property.
        doc.write_protection.read_only_recommended = True
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentProperties.Security.ReadOnlyRecommended.docx')
        self.assertEqual(aw.properties.DocumentSecurity.READ_ONLY_RECOMMENDED, aw.Document(file_name=ARTIFACTS_DIR + 'DocumentProperties.Security.ReadOnlyRecommended.docx').built_in_document_properties.security)
        # Write-protect a document, and then verify its security level.
        doc = aw.Document()
        self.assertFalse(doc.write_protection.is_write_protected)
        doc.write_protection.set_password('MyPassword')
        self.assertTrue(doc.write_protection.validate_password('MyPassword'))
        self.assertTrue(doc.write_protection.is_write_protected)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentProperties.Security.ReadOnlyEnforced.docx')
        self.assertEqual(aw.properties.DocumentSecurity.READ_ONLY_ENFORCED, aw.Document(file_name=ARTIFACTS_DIR + 'DocumentProperties.Security.ReadOnlyEnforced.docx').built_in_document_properties.security)
        # "Security" is a descriptive property. We can edit its value manually.
        doc = aw.Document()
        doc.protect(type=aw.ProtectionType.ALLOW_ONLY_COMMENTS, password='MyPassword')
        doc.built_in_document_properties.security = aw.properties.DocumentSecurity.READ_ONLY_EXCEPT_ANNOTATIONS
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentProperties.Security.ReadOnlyExceptAnnotations.docx')
        self.assertEqual(aw.properties.DocumentSecurity.READ_ONLY_EXCEPT_ANNOTATIONS, aw.Document(file_name=ARTIFACTS_DIR + 'DocumentProperties.Security.ReadOnlyExceptAnnotations.docx').built_in_document_properties.security)
        #ExEnd

    def test_link_custom_document_properties_to_bookmark(self):
        #ExStart
        #ExFor:CustomDocumentProperties.add_link_to_content(str,str)
        #ExFor:DocumentProperty.is_link_to_content
        #ExFor:DocumentProperty.link_source
        #ExSummary:Shows how to link a custom document property to a bookmark.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.start_bookmark('MyBookmark')
        builder.write('Hello world!')
        builder.end_bookmark('MyBookmark')
        # Link a new custom property to a bookmark. The value of this property
        # will be the contents of the bookmark that it references in the "LinkSource" member.
        custom_properties = doc.custom_document_properties
        custom_property = custom_properties.add_link_to_content('Bookmark', 'MyBookmark')
        self.assertEqual(True, custom_property.is_link_to_content)
        self.assertEqual('MyBookmark', custom_property.link_source)
        self.assertEqual('Hello world!', custom_property.value)
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentProperties.LinkCustomDocumentPropertiesToBookmark.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'DocumentProperties.LinkCustomDocumentPropertiesToBookmark.docx')
        custom_property = doc.custom_document_properties.get_by_name('Bookmark')
        self.assertEqual(True, custom_property.is_link_to_content)
        self.assertEqual('MyBookmark', custom_property.link_source)
        self.assertEqual('Hello world!', custom_property.value)

    def test_document_property_collection(self):
        #ExStart
        #ExFor:CustomDocumentProperties.add(str,str)
        #ExFor:CustomDocumentProperties.add(str,bool)
        #ExFor:CustomDocumentProperties.add(str,int)
        #ExFor:CustomDocumentProperties.add(str,datetime)
        #ExFor:CustomDocumentProperties.add(str,float)
        #ExFor:DocumentProperty.type
        #ExFor:DocumentPropertyCollection
        #ExFor:DocumentPropertyCollection.clear
        #ExFor:DocumentPropertyCollection.contains(str)
        #ExFor:DocumentPropertyCollection.index_of(str)
        #ExFor:DocumentPropertyCollection.remove_at(int)
        #ExFor:DocumentPropertyCollection.remove
        #ExFor:PropertyType
        #ExSummary:Shows how to work with a document's custom properties.
        import datetime
        import aspose.words as aw
        from api_example_base import ApiExampleBase, ARTIFACTS_DIR
        doc = aw.Document()
        properties = doc.custom_document_properties
        self.assertEqual(0, properties.count)
        # Custom document properties are key-value pairs that we can add to the document.
        properties.add(name='Authorized', value=True)
        properties.add(name='Authorized By', value='John Doe')
        properties.add(name='Authorized Date', value=datetime.date.today())
        properties.add(name='Authorized Revision', value=doc.built_in_document_properties.revision_number)
        properties.add(name='Authorized Amount', value=123.45)
        # The collection sorts the custom properties in alphabetic order.
        self.assertEqual(1, properties.index_of('Authorized Amount'))
        self.assertEqual(5, properties.count)
        # Print every custom property in the document.
        for prop in properties:
            print(f'Name: "{prop.name}"\n\tType: "{prop.type}"\n\tValue: "{prop.value}"')
        # Display the value of a custom property using a DOCPROPERTY field.
        builder = aw.DocumentBuilder(doc=doc)
        field = builder.insert_field(field_code=' DOCPROPERTY "Authorized By"').as_field_doc_property()
        field.update()
        self.assertEqual('John Doe', field.result)
        # We can find these custom properties in Microsoft Word via "File" -> "Properties" > "Advanced Properties" > "Custom".
        doc.save(file_name=ARTIFACTS_DIR + 'DocumentProperties.DocumentPropertyCollection.docx')
        # Below are three ways or removing custom properties from a document.
        # 1 -  Remove by index:
        properties.remove_at(1)
        self.assertFalse(properties.contains('Authorized Amount'))
        self.assertEqual(4, properties.count)
        # 2 -  Remove by name:
        properties.remove('Authorized Revision')
        self.assertFalse(properties.contains('Authorized Revision'))
        self.assertEqual(3, properties.count)
        # 3 -  Empty the entire collection at once:
        properties.clear()
        self.assertEqual(0, properties.count)
        #ExEnd

    def test_extended_properties(self):
        #ExStart:ExtendedProperties
        #ExFor:BuiltInDocumentProperties.scale_crop
        #ExFor:BuiltInDocumentProperties.shared_document
        #ExFor:BuiltInDocumentProperties.hyperlinks_changed
        #ExSummary:Shows how to get extended properties.
        doc = aw.Document(file_name=MY_DIR + 'Extended properties.docx')
        self.assertTrue(doc.built_in_document_properties.scale_crop)
        self.assertTrue(doc.built_in_document_properties.shared_document)
        self.assertTrue(doc.built_in_document_properties.hyperlinks_changed)
        #ExEnd:ExtendedProperties
    #ExStart
    #ExFor:BuiltInDocumentProperties.bytes
    #ExFor:BuiltInDocumentProperties.characters
    #ExFor:BuiltInDocumentProperties.characters_with_spaces
    #ExFor:BuiltInDocumentProperties.content_status
    #ExFor:BuiltInDocumentProperties.content_type
    #ExFor:BuiltInDocumentProperties.lines
    #ExFor:BuiltInDocumentProperties.links_up_to_date
    #ExFor:BuiltInDocumentProperties.pages
    #ExFor:BuiltInDocumentProperties.paragraphs
    #ExFor:BuiltInDocumentProperties.words
    #ExSummary:Shows how to work with document properties in the "Content" category (LineCounter).

    class LineCounter:

        def __init__(self, doc):
            self.m_scanning_line_for_real_text = None
            self.m_line_count = None
            self.m_layout_enumerator = aw.layout.LayoutEnumerator(doc)
            self._count_lines()

        def get_line_count(self):
            return self.m_line_count

        def _count_lines(self):
            while true:
                if self.m_layout_enumerator.type == aw.layout.LayoutEntityType.LINE:
                    self.m_scanning_line_for_real_text = True
                if self.m_layout_enumerator.move_first_child():
                    if self.m_scanning_line_for_real_text and self.m_layout_enumerator.kind.startswith('TEXT'):
                        self.m_line_count += 1
                        self.m_scanning_line_for_real_text = False
                    self._count_lines()
                    self.m_layout_enumerator.move_parent()
                if self.m_layout_enumerator.move_next():
                    break
    #ExEnd

    def test_custom(self):
        #ExStart
        #ExFor:BuiltInDocumentProperties.__getitem__(str)
        #ExFor:CustomDocumentProperties
        #ExFor:DocumentProperty.__str__
        #ExFor:DocumentPropertyCollection.count
        #ExFor:DocumentPropertyCollection.__getitem__(int)
        #ExSummary:Shows how to work with custom document properties.
        doc = aw.Document(MY_DIR + 'Properties.docx')
        # Every document contains a collection of custom properties, which, like the built-in properties, are key-value pairs.
        # The document has a fixed list of built-in properties. The user creates all of the custom properties.
        self.assertEqual('Value of custom document property', str(doc.custom_document_properties.get_by_name('CustomProperty')))
        doc.custom_document_properties.add('CustomProperty2', 'Value of custom document property #2')
        print('Custom Properties:')
        for custom_document_property in doc.custom_document_properties:
            print(custom_document_property.name)
            print(f'\tType:\t{custom_document_property.type}')
            print(f'\tValue:\t"{custom_document_property.value}"')
        #ExEnd
        self.assertEqual(2, doc.custom_document_properties.count)