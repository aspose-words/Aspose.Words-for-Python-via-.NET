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

class ExSmartTag(ApiExampleBase):

    #ExStart
    #ExFor:CompositeNode.remove_smart_tags
    #ExFor:CustomXmlProperty
    #ExFor:CustomXmlProperty.__init__(str,str,str)
    #ExFor:CustomXmlProperty.name
    #ExFor:CustomXmlProperty.value
    #ExFor:SmartTag
    #ExFor:SmartTag.__init__(DocumentBase)
    #ExFor:SmartTag.accept(DocumentVisitor)
    #ExFor:SmartTag.element
    #ExFor:SmartTag.properties
    #ExFor:SmartTag.uri
    #ExSummary:Shows how to create smart tags.
    def test_create(self):

        doc = aw.Document()

        # A smart tag appears in a document with Microsoft Word recognizes a part of its text as some form of data,
        # such as a name, date, or address, and converts it to a hyperlink that displays a purple dotted underline.
        smart_tag = aw.markup.SmartTag(doc)

        # Smart tags are composite nodes that contain their recognized text in its entirety.
        # Add contents to this smart tag manually.
        smart_tag.append_child(aw.Run(doc, "May 29, 2019"))

        # Microsoft Word may recognize the above contents as being a date.
        # Smart tags use the "Element" property to reflect the type of data they contain.
        smart_tag.element = "date"

        # Some smart tag types process their contents further into custom XML properties.
        smart_tag.properties.add(aw.markup.CustomXmlProperty("Day", "", "29"))
        smart_tag.properties.add(aw.markup.CustomXmlProperty("Month", "", "5"))
        smart_tag.properties.add(aw.markup.CustomXmlProperty("Year", "", "2019"))

        # Set the smart tag's URI to the default value.
        smart_tag.uri = "urn:schemas-microsoft-com:office:smarttags"

        doc.first_section.body.first_paragraph.append_child(smart_tag)
        doc.first_section.body.first_paragraph.append_child(aw.Run(doc, " is a date. "))

        # Create another smart tag for a stock ticker.
        smart_tag = aw.markup.SmartTag(doc)
        smart_tag.element = "stockticker"
        smart_tag.uri = "urn:schemas-microsoft-com:office:smarttags"

        smart_tag.append_child(aw.Run(doc, "MSFT"))

        doc.first_section.body.first_paragraph.append_child(smart_tag)
        doc.first_section.body.first_paragraph.append_child(aw.Run(doc, " is a stock ticker."))

        # Print all the smart tags in our document using a document visitor.
        #doc.accept(ExSmartTag.SmartTagPrinter())

        # Older versions of Microsoft Word support smart tags.
        doc.save(ARTIFACTS_DIR + "SmartTag.create.doc")

        # Use the "remove_smart_tags" method to remove all smart tags from a document.
        self.assertEqual(2, doc.get_child_nodes(aw.NodeType.SMART_TAG, True).count)

        doc.remove_smart_tags()

        self.assertEqual(0, doc.get_child_nodes(aw.NodeType.SMART_TAG, True).count)
        self._test_create(aw.Document(ARTIFACTS_DIR + "SmartTag.create.doc")) #ExSkip

    #class SmartTagPrinter(aw.DocumentVisitor):
    #    """Prints visited smart tags and their contents."""

    #    def visit_smart_tag_start(smart_tag: aw.markup.SmartTag) -> aw.VisitorAction:
    #        """Called when a SmartTag node is encountered in the document."""

    #        print(f"Smart tag type: {smart_tag.element}")
    #        return aw.VisitorAction.CONTINUE

    #    def visit_smart_tag_end(smart_tag: aw.markup.SmartTag) -> aw.VisitorAction:
    #        """Called when the visiting of a SmartTag node is ended."""

    #        print(f"\tContents: \"{smart_tag.to_string(aw.SaveFormat.TEXT)}\"")

    #        if smart_tag.properties.count == 0:
    #            print("\tContains no properties")
    #        else:

    #            print("\tProperties: ", end="")
    #            properties = string[smart_tag.properties.count]
    #            index = 0

    #            for cxp in smart_tag.properties:
    #                properties[index] = f'"{cxp.Name}" = "{cxp.Value}"'
    #                index += 1

    #            print("".join(", ", properties))

    #        return aw.VisitorAction.CONTINUE

    #ExEnd

    def _test_create(self, doc: aw.Document):

        smart_tag = doc.get_child(aw.NodeType.SMART_TAG, 0, True).as_smart_tag()

        self.assertEqual("date", smart_tag.element)
        self.assertEqual("May 29, 2019", smart_tag.get_text())
        self.assertEqual("urn:schemas-microsoft-com:office:smarttags", smart_tag.uri)

        self.assertEqual("Day", smart_tag.properties[0].name)
        self.assertEqual("", smart_tag.properties[0].uri)
        self.assertEqual("29", smart_tag.properties[0].value)
        self.assertEqual("Month", smart_tag.properties[1].name)
        self.assertEqual("", smart_tag.properties[1].uri)
        self.assertEqual("5", smart_tag.properties[1].value)
        self.assertEqual("Year", smart_tag.properties[2].name)
        self.assertEqual("", smart_tag.properties[2].uri)
        self.assertEqual("2019", smart_tag.properties[2].value)

        smart_tag = doc.get_child(aw.NodeType.SMART_TAG, 1, True).as_smart_tag()

        self.assertEqual("stockticker", smart_tag.element)
        self.assertEqual("MSFT", smart_tag.get_text())
        self.assertEqual("urn:schemas-microsoft-com:office:smarttags", smart_tag.uri)
        self.assertEqual(0, smart_tag.properties.count)

    def test_properties(self):

        #ExStart
        #ExFor:CustomXmlProperty.uri
        #ExFor:CustomXmlPropertyCollection
        #ExFor:CustomXmlPropertyCollection.add(CustomXmlProperty)
        #ExFor:CustomXmlPropertyCollection.clear
        #ExFor:CustomXmlPropertyCollection.contains(str)
        #ExFor:CustomXmlPropertyCollection.count
        #ExFor:CustomXmlPropertyCollection.__iter__
        #ExFor:CustomXmlPropertyCollection.index_of_key(str)
        #ExFor:CustomXmlPropertyCollection.__getitem__(int)
        #ExFor:CustomXmlPropertyCollection.__getitem__(str)
        #ExFor:CustomXmlPropertyCollection.remove(str)
        #ExFor:CustomXmlPropertyCollection.remove_at(int)
        #ExSummary:Shows how to work with smart tag properties to get in depth information about smart tags.
        doc = aw.Document(MY_DIR + "Smart tags.doc")

        # A smart tag appears in a document with Microsoft Word recognizes a part of its text as some form of data,
        # such as a name, date, or address, and converts it to a hyperlink that displays a purple dotted underline.
        # In Word 2003, we can enable smart tags via "Tools" -> "AutoCorrect options..." -> "SmartTags".
        # In our input document, there are three objects that Microsoft Word registered as smart tags.
        # Smart tags may be nested, so this collection contains more.
        smart_tags = [node.as_smart_tag() for node in doc.get_child_nodes(aw.NodeType.SMART_TAG, True)]

        self.assertEqual(8, len(smart_tags))

        # The "properties" member of a smart tag contains its metadata, which will be different for each type of smart tag.
        # The properties of a "date"-type smart tag contain its year, month, and day.
        properties = smart_tags[7].properties

        self.assertEqual(4, properties.count)

        for prop in properties:
            print(f"Property name: {prop.name}, value: {prop.value}")
            self.assertEqual("", prop.uri)

        # We can also access the properties in various ways, such as a key-value pair.
        self.assertTrue(properties.contains("Day"))
        self.assertEqual("22", properties.get_by_name("Day").value)
        self.assertEqual("2003", properties[2].value)
        self.assertEqual(1, properties.index_of_key("Month"))

        # Below are three ways of removing elements from the properties collection.
        # 1 -  Remove by index:
        properties.remove_at(3)

        self.assertEqual(3, properties.count)

        # 2 -  Remove by name:
        properties.remove("Year")

        self.assertEqual(2, properties.count)

        # 3 -  Clear the entire collection at once:
        properties.clear()

        self.assertEqual(0, properties.count)
        #ExEnd
