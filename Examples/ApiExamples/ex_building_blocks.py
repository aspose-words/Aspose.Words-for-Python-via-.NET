# Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.

import unittest
import io
import uuid
from typing import Dict

import aspose.words as aw
import aspose.pydrawing as drawing

from api_example_base import ApiExampleBase, my_dir, artifacts_dir

MY_DIR = my_dir
ARTIFACTS_DIR = artifacts_dir

class ExBuildingBlocks(ApiExampleBase):

    ##ExStart
    ##ExFor:Document.glossary_document
    ##ExFor:BuildingBlocks.building_block
    ##ExFor:BuildingBlocks.BuildingBlock.__init__(GlossaryDocument)
    ##ExFor:BuildingBlocks.BuildingBlock.accept(DocumentVisitor)
    ##ExFor:BuildingBlocks.BuildingBlock.behavior
    ##ExFor:BuildingBlocks.BuildingBlock.category
    ##ExFor:BuildingBlocks.BuildingBlock.description
    ##ExFor:BuildingBlocks.BuildingBlock.first_section
    ##ExFor:BuildingBlocks.BuildingBlock.gallery
    ##ExFor:BuildingBlocks.BuildingBlock.guid
    ##ExFor:BuildingBlocks.BuildingBlock.last_section
    ##ExFor:BuildingBlocks.BuildingBlock.name
    ##ExFor:BuildingBlocks.BuildingBlock.sections
    ##ExFor:BuildingBlocks.BuildingBlock.type
    ##ExFor:BuildingBlocks.building_block_behavior
    ##ExFor:BuildingBlocks.building_block_type
    ##ExSummary:Shows how to add a custom building block to a document.
    #def test_create_and_insert(self):

    #    # A document's glossary document stores building blocks.
    #    doc = aw.Document()
    #    glossary_doc = aw.buildingblocks.GlossaryDocument()
    #    doc.glossary_document = glossary_doc

    #    # Create a building block, name it, and then add it to the glossary document.
    #    block = aw.buildingblocks.BuildingBlock(glossary_doc)
    #    block.name = "Custom Block"

    #    glossary_doc.append_child(block)

    #    # All new building block GUIDs have the same zero value by default, and we can give them a new unique value.
    #    self.assertEqual(uuid.UUID(), block.guid)

    #    block.guid = uuid.uuid4()

    #    # The following properties categorize building blocks
    #    # in the menu we can access in Microsoft Word via "Insert" -> "Quick Parts" -> "Building Blocks Organizer".
    #    self.assertEqual("(Empty Category)", block.category)
    #    self.assertEqual(aw.buildingblocks.BuildingBlockType.NONE, block.type)
    #    self.assertEqual(aw.buildingblocks.BuildingBlockGallery.ALL, block.gallery)
    #    self.assertEqual(aw.buildingblocks.BuildingBlockBehavior.CONTENT, block.behavior)

    #    # Before we can add this building block to our document, we will need to give it some contents,
    #    # which we will do using a document visitor. This visitor will also set a category, gallery, and behavior.
    #    visitor = ExBuildingBlocks.BuildingBlockVisitor(glossary_doc)
    #    block.accept(visitor)

    #    # We can access the block that we just made from the glossary document.
    #    custom_block = glossary_doc.get_building_block(aw.buildingblocks.BuildingBlockGallery.QUICK_PARTS,
    #        "My custom building blocks", "Custom Block")

    #    # The block itself is a section that contains the text.
    #    self.assertEqual(f"Text inside {custom_block.Name}\f", custom_block.first_section.body.first_paragraph.get_text())
    #    self.assertEqual(custom_block.first_section, custom_block.last_section)
    #    uuid.UUID(custom_block.guid) #ExSkip
    #    self.assertEqual("My custom building blocks", custom_block.category) #ExSkip
    #    self.assertEqual(aw.buildingblocks.BuildingBlockType.NONE, custom_block.type) #ExSkip
    #    self.assertEqual(aw.buildingblocks.BuildingBlockGallery.QUICK_PARTS, custom_block.gallery) #ExSkip
    #    self.assertEqual(aw.buildingblocks.BuildingBlockBehavior.PARAGRAPH, custom_block.behavior) #ExSkip

    #    # Now, we can insert it into the document as a new section.
    #    doc.append_child(doc.import_node(custom_block.first_section, True))

    #    # We can also find it in Microsoft Word's Building Blocks Organizer and place it manually.
    #    doc.save(ARTIFACTS_DIR + "BuildingBlocks.create_and_insert.dotx")

    #class BuildingBlockVisitor(aw.DocumentVisitor):
    #    """Sets up a visited building block to be inserted into the document as a quick part and adds text to its contents."""

    #    def __init__(self, owner_glossary_doc: aw.buildingblocks.GlossaryDocument):

    #        self.builder = io.StringIO()
    #        self.glossary_doc = owner_glossary_doc

    #    def visit_building_block_start(self, block: aw.buildingblocks.BuildingBlock) -> aw.VisitorAction:

    #        # Configure the building block as a quick part, and add properties used by Building Blocks Organizer.
    #        block.behavior = aw.buildingblocks.BuildingBlockBehavior.PARAGRAPH
    #        block.category = "My custom building blocks"
    #        block.description = "Using this block in the Quick Parts section of word will place its contents at the cursor."
    #        block.gallery = aw.buildingblocks.BuildingBlockGallery.QUICK_PARTS

    #        # Add a section with text.
    #        # Inserting the block into the document will append this section with its child nodes at the location.
    #        section = aw.Section(self.glossary_doc)
    #        block.append_child(section)
    #        block.first_section.ensure_minimum()

    #        run = aw.Run(self.glossary_doc, "Text inside " + block.name)
    #        block.first_section.body.first_paragraph.append_child(run)

    #        return aw.VisitorAction.CONTINUE

    #    def visit_building_block_end(self, block: aw.buildingblocks.BuildingBlock) -> aw.VisitorAction:

    #        self.builder.append("Visited " + block.name + "\r\n")
    #        return aw.VisitorAction.CONTINUE

    ##ExEnd

    ##ExStart
    ##ExFor:BuildingBlocks.glossary_document
    ##ExFor:BuildingBlocks.GlossaryDocument.accept(DocumentVisitor)
    ##ExFor:BuildingBlocks.GlossaryDocument.building_blocks
    ##ExFor:BuildingBlocks.GlossaryDocument.first_building_block
    ##ExFor:BuildingBlocks.GlossaryDocument.get_building_block(BuildingBlocks.BuildingBlockGallery,str,str)
    ##ExFor:BuildingBlocks.GlossaryDocument.last_building_block
    ##ExFor:BuildingBlocks.building_block_collection
    ##ExFor:BuildingBlocks.BuildingBlockCollection.__getitem__(int)
    ##ExFor:BuildingBlocks.BuildingBlockCollection.to_array
    ##ExFor:BuildingBlocks.building_block_gallery
    ##ExFor:DocumentVisitor.visit_building_block_end(BuildingBlock)
    ##ExFor:DocumentVisitor.visit_building_block_start(BuildingBlock)
    ##ExFor:DocumentVisitor.visit_glossary_document_end(GlossaryDocument)
    ##ExFor:DocumentVisitor.visit_glossary_document_start(GlossaryDocument)
    ##ExSummary:Shows ways of accessing building blocks in a glossary document.
    #def test_glossary_document(self):

    #    doc = aw.Document()
    #    glossary_doc = aw.buildingblocks.GlossaryDocument()

    #    for i in range(1, 6):
    #        block = aw.buildingblocks.BuildingBlock(glossary_doc)
    #        block.name = f"Block {i}"
    #        glossary_doc.append_child(block)
        
    #    self.assertEqual(5, glossary_doc.building_blocks.count)

    #    doc.glossary_document = glossary_doc

    #    # There are various ways of accessing building blocks.
    #    # 1 -  Get the first/last building blocks in the collection:
    #    self.assertEqual("Block 1", glossary_doc.first_building_block.name)
    #    self.assertEqual("Block 5", glossary_doc.last_building_block.name)

    #    # 2 -  Get a building block by index:
    #    self.assertEqual("Block 2", glossary_doc.building_blocks[1].name)
    #    self.assertEqual("Block 3", glossary_doc.building_blocks.to_array()[2].name)

    #    # 3 -  Get the first building block that matches a gallery, name and category:
    #    self.assertEqual("Block 4",
    #        glossary_doc.get_building_block(aw.buildingblocks.BuildingBlockGallery.ALL, "(Empty Category)", "Block 4").name)

    #    # We will do that using a custom visitor,
    #    # which will give every BuildingBlock in the GlossaryDocument a unique GUID
    #    visitor = ExBuldingBlocks.GlossaryDocVisitor()
    #    glossary_doc.accept(visitor)
    #    self.assertEqual(5, visitor.get_dictionary().count) #ExSkip

    #    print(visitor.get_text())

    #    # In Microsoft Word, we can access the building blocks via "Insert" -> "Quick Parts" -> "Building Blocks Organizer".
    #    doc.save(ARTIFACTS_DIR + "BuildingBlocks.glossary_document.dotx")

    #class GlossaryDocVisitor(aw.DocumentVisitor):
    #    """Gives each building block in a visited glossary document a unique GUID.
    #    Stores the GUID-building block pairs in a dictionary."""

    #    def __init__(self):

    #        self.blocks_by_guid: Dict[uuid.UUID, aw.buildingblocks.BuildingBlock] = {}
    #        self.builder = io.StringIO()

    #    def get_text():

    #        return self.builder.getvalue()

    #    def get_dictionary(self) -> Dict[uuid.UUID, aw.buildingblocks.BuildingBlock]:

    #        return self.blocks_by_guid

    #    def visit_glossary_document_start(self, glossary: aw.buildingblocks.GlossaryDocument) -> aw.VisitorAction:

    #        self.builder.write("Glossary document found!\n")
    #        return aw.VisitorAction.CONTINUE

    #    def visit_glossary_document_end(self, glossary: aw.buildingblocks.GlossaryDocument) -> aw.VisitorAction:

    #        self.builder.write("Reached end of glossary!\n")
    #        self.builder.write("BuildingBlocks found: " + self.blocks_by_guid.count + "\n")
    #        return aw.VisitorAction.CONTINUE

    #    def visit_building_block_start(self, block: aw.buildingblocks.BuildingBlock) -> aw.VisitorAction:

    #        self.assertEqual("00000000-0000-0000-0000-000000000000", block.guid.to_string()) #ExSkip
    #        block.guid = uuid.uuid4()
    #        self.blocks_by_guid.add(block.guid, block)
    #        return aw.VisitorAction.CONTINUE

    #    def visit_building_block_end(self, block: aw.buildingblocks.BuildingBlock) -> aw.VisitorAction:

    #        self.builder.write("\tVisited block \"" + block.name + "\"\n")
    #        self.builder.write("\t Type: " + block.type + "\n")
    #        self.builder.write("\t Gallery: " + block.gallery + "\n")
    #        self.builder.write("\t Behavior: " + block.behavior + "\n")
    #        self.builder.write("\t Description: " + block.description + "\n")

    #        return aw.VisitorAction.CONTINUE

    ##ExEnd
