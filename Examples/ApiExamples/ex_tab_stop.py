# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
import system_helper
import test_util
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR

class ExTabStop(ApiExampleBase):

    def test_add_tab_stops(self):
        #ExStart
        #ExFor:TabStopCollection.add(TabStop)
        #ExFor:TabStopCollection.add(float,TabAlignment,TabLeader)
        #ExSummary:Shows how to add custom tab stops to a document.
        doc = aw.Document()
        paragraph = doc.get_child(aw.NodeType.PARAGRAPH, 0, True).as_paragraph()
        # Below are two ways of adding tab stops to a paragraph's collection of tab stops via the "ParagraphFormat" property.
        # 1 -  Create a "TabStop" object, and then add it to the collection:
        tab_stop = aw.TabStop(position=aw.ConvertUtil.inch_to_point(3), alignment=aw.TabAlignment.LEFT, leader=aw.TabLeader.DASHES)
        paragraph.paragraph_format.tab_stops.add(tab_stop=tab_stop)
        # 2 -  Pass the values for properties of a new tab stop to the "Add" method:
        paragraph.paragraph_format.tab_stops.add(position=aw.ConvertUtil.millimeter_to_point(100), alignment=aw.TabAlignment.LEFT, leader=aw.TabLeader.DASHES)
        # Add tab stops at 5 cm to all paragraphs.
        for para in filter(lambda a: a is not None, map(lambda b: system_helper.linq.Enumerable.of_type(lambda x: x.as_paragraph(), b), list(doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)))):
            para.paragraph_format.tab_stops.add(position=aw.ConvertUtil.millimeter_to_point(50), alignment=aw.TabAlignment.LEFT, leader=aw.TabLeader.DASHES)
        # Every "tab" character takes the builder's cursor to the location of the next tab stop.
        builder = aw.DocumentBuilder(doc=doc)
        builder.writeln('Start\tTab 1\tTab 2\tTab 3\tTab 4')
        doc.save(file_name=ARTIFACTS_DIR + 'TabStopCollection.AddTabStops.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'TabStopCollection.AddTabStops.docx')
        tab_stops = doc.first_section.body.paragraphs[0].paragraph_format.tab_stops
        test_util.TestUtil.verify_tab_stop(141.75, aw.TabAlignment.LEFT, aw.TabLeader.DASHES, False, tab_stops[0])
        test_util.TestUtil.verify_tab_stop(216, aw.TabAlignment.LEFT, aw.TabLeader.DASHES, False, tab_stops[1])
        test_util.TestUtil.verify_tab_stop(283.45, aw.TabAlignment.LEFT, aw.TabLeader.DASHES, False, tab_stops[2])

    def test_remove_by_index(self):
        #ExStart
        #ExFor:TabStopCollection.remove_by_index
        #ExSummary:Shows how to select a tab stop in a document by its index and remove it.
        doc = aw.Document()
        tab_stops = doc.first_section.body.paragraphs[0].paragraph_format.tab_stops
        tab_stops.add(position=aw.ConvertUtil.millimeter_to_point(30), alignment=aw.TabAlignment.LEFT, leader=aw.TabLeader.DASHES)
        tab_stops.add(position=aw.ConvertUtil.millimeter_to_point(60), alignment=aw.TabAlignment.LEFT, leader=aw.TabLeader.DASHES)
        self.assertEqual(2, tab_stops.count)
        # Remove the first tab stop.
        tab_stops.remove_by_index(0)
        self.assertEqual(1, tab_stops.count)
        doc.save(file_name=ARTIFACTS_DIR + 'TabStopCollection.RemoveByIndex.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'TabStopCollection.RemoveByIndex.docx')
        test_util.TestUtil.verify_tab_stop(170.1, aw.TabAlignment.LEFT, aw.TabLeader.DASHES, False, doc.first_section.body.paragraphs[0].paragraph_format.tab_stops[0])

    def test_get_position_by_index(self):
        #ExStart
        #ExFor:TabStopCollection.get_position_by_index
        #ExSummary:Shows how to find a tab, stop by its index and verify its position.
        doc = aw.Document()
        tab_stops = doc.first_section.body.paragraphs[0].paragraph_format.tab_stops
        tab_stops.add(position=aw.ConvertUtil.millimeter_to_point(30), alignment=aw.TabAlignment.LEFT, leader=aw.TabLeader.DASHES)
        tab_stops.add(position=aw.ConvertUtil.millimeter_to_point(60), alignment=aw.TabAlignment.LEFT, leader=aw.TabLeader.DASHES)
        # Verify the position of the second tab stop in the collection.
        self.assertAlmostEqual(aw.ConvertUtil.millimeter_to_point(60), tab_stops.get_position_by_index(1), delta=0.1)
        #ExEnd

    def test_get_index_by_position(self):
        #ExStart
        #ExFor:TabStopCollection.get_index_by_position
        #ExSummary:Shows how to look up a position to see if a tab stop exists there and obtain its index.
        doc = aw.Document()
        tab_stops = doc.first_section.body.paragraphs[0].paragraph_format.tab_stops
        # Add a tab stop at a position of 30mm.
        tab_stops.add(position=aw.ConvertUtil.millimeter_to_point(30), alignment=aw.TabAlignment.LEFT, leader=aw.TabLeader.DASHES)
        # A result of "0" returned by "GetIndexByPosition" confirms that a tab stop
        # at 30mm exists in this collection, and it is at index 0.
        self.assertEqual(0, tab_stops.get_index_by_position(aw.ConvertUtil.millimeter_to_point(30)))
        # A "-1" returned by "GetIndexByPosition" confirms that
        # there is no tab stop in this collection with a position of 60mm.
        self.assertEqual(-1, tab_stops.get_index_by_position(aw.ConvertUtil.millimeter_to_point(60)))
        #ExEnd

    def test_tab_stop_collection(self):
        #ExStart
        #ExFor:TabStop.__init__(float)
        #ExFor:TabStop.__init__(float,TabAlignment,TabLeader)
        #ExFor:TabStop.__eq__(TabStop)
        #ExFor:TabStop.is_clear
        #ExFor:TabStopCollection
        #ExFor:TabStopCollection.after(float)
        #ExFor:TabStopCollection.before(float)
        #ExFor:TabStopCollection.clear
        #ExFor:TabStopCollection.count
        #ExFor:TabStopCollection.__eq__(TabStopCollection)
        #ExFor:TabStopCollection.__eq__(object)
        #ExFor:TabStopCollection.__hash__
        #ExFor:TabStopCollection.__getitem__(float)
        #ExFor:TabStopCollection.__getitem__(int)
        #ExSummary:Shows how to work with a document's collection of tab stops.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        tab_stops = builder.paragraph_format.tab_stops
        # 72 points is one "inch" on the Microsoft Word tab stop ruler.
        tab_stops.add(tab_stop=aw.TabStop(position=72))
        tab_stops.add(tab_stop=aw.TabStop(position=432, alignment=aw.TabAlignment.RIGHT, leader=aw.TabLeader.DASHES))
        self.assertEqual(2, tab_stops.count)
        self.assertFalse(tab_stops[0].is_clear)
        self.assertFalse(tab_stops[0].equals(tab_stops[1]))
        # Every "tab" character takes the builder's cursor to the location of the next tab stop.
        builder.writeln('Start\tTab 1\tTab 2')
        paragraphs = doc.first_section.body.paragraphs
        self.assertEqual(2, paragraphs.count)
        # Each paragraph gets its tab stop collection, which clones its values from the document builder's tab stop collection.
        self.assertEqual(paragraphs[0].paragraph_format.tab_stops, paragraphs[1].paragraph_format.tab_stops)
        # A tab stop collection can point us to TabStops before and after certain positions.
        self.assertEqual(72, tab_stops.before(100).position)
        self.assertEqual(432, tab_stops.after(100).position)
        # We can clear a paragraph's tab stop collection to revert to the default tabbing behavior.
        paragraphs[1].paragraph_format.tab_stops.clear()
        self.assertEqual(0, paragraphs[1].paragraph_format.tab_stops.count)
        doc.save(file_name=ARTIFACTS_DIR + 'TabStopCollection.TabStopCollection.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'TabStopCollection.TabStopCollection.docx')
        tab_stops = doc.first_section.body.paragraphs[0].paragraph_format.tab_stops
        self.assertEqual(2, tab_stops.count)
        test_util.TestUtil.verify_tab_stop(72, aw.TabAlignment.LEFT, aw.TabLeader.NONE, False, tab_stops[0])
        test_util.TestUtil.verify_tab_stop(432, aw.TabAlignment.RIGHT, aw.TabLeader.DASHES, False, tab_stops[1])
        tab_stops = doc.first_section.body.paragraphs[1].paragraph_format.tab_stops
        self.assertEqual(0, tab_stops.count)