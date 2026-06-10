# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
import aspose.words.markup
import system_helper
import unittest
from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR

class ExSmartTag(ApiExampleBase):

    def test_properties(self):
        from api_example_base import ApiExampleBase, MY_DIR, ARTIFACTS_DIR, GOLDS_DIR, TEMP_DIR, IMAGE_DIR, FONTS_DIR
        import aspose.words as aw
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
        doc = aw.Document(file_name=MY_DIR + 'Smart tags.doc')
        # A smart tag appears in a document with Microsoft Word recognizes a part of its text as some form of data,
        # such as a name, date, or address, and converts it to a hyperlink that displays a purple dotted underline.
        # In Word 2003, we can enable smart tags via "Tools" -> "AutoCorrect options..." -> "SmartTags".
        # In our input document, there are three objects that Microsoft Word registered as smart tags.
        # Smart tags may be nested, so this collection contains more.
        smart_tags = list(filter(lambda a: a is not None, map(lambda b: system_helper.linq.Enumerable.of_type(lambda x: x.as_smart_tag(), b), list(doc.get_child_nodes(aw.NodeType.SMART_TAG, True)))))
        self.assertEqual(8, len(smart_tags))
        # The "Properties" member of a smart tag contains its metadata, which will be different for each type of smart tag.
        # The properties of a "date"-type smart tag contain its year, month, and day.
        properties = smart_tags[7].properties
        self.assertEqual(4, properties.count)
        for current in properties:
            print(f'Property name: {current.name}, value: {current.value}')
            self.assertEqual('', current.uri)
        # We can also access the properties in various ways, such as a key-value pair.
        self.assertTrue(properties.contains('Day'))
        self.assertEqual('22', properties.get_by_name('Day').value)
        self.assertEqual('2003', properties[2].value)
        self.assertEqual(1, properties.index_of_key('Month'))
        # Below are three ways of removing elements from the properties collection.
        # 1 -  Remove by index:
        properties.remove_at(3)
        self.assertEqual(3, properties.count)
        # 2 -  Remove by name:
        properties.remove('Year')
        self.assertEqual(2, properties.count)
        # 3 -  Clear the entire collection at once:
        properties.clear()
        self.assertEqual(0, properties.count)
        #ExEnd