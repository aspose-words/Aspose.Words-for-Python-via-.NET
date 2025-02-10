# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
from datetime import date
import aspose.words as aw
import aspose.words.replacing
import aspose.words.saving
import datetime
import system_helper
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR, IMAGE_DIR

class ExHeaderFooter(ApiExampleBase):

    def test_create(self):
        #ExStart
        #ExFor:HeaderFooter
        #ExFor:HeaderFooter.__init__(DocumentBase,HeaderFooterType)
        #ExFor:HeaderFooter.header_footer_type
        #ExFor:HeaderFooter.is_header
        #ExFor:HeaderFooterCollection
        #ExFor:Paragraph.is_end_of_header_footer
        #ExFor:Paragraph.parent_section
        #ExFor:Paragraph.parent_story
        #ExFor:Story.append_paragraph
        #ExSummary:Shows how to create a header and a footer.
        doc = aw.Document()
        # Create a header and append a paragraph to it. The text in that paragraph
        # will appear at the top of every page of this section, above the main body text.
        header = aw.HeaderFooter(doc, aw.HeaderFooterType.HEADER_PRIMARY)
        doc.first_section.headers_footers.add(header)
        para = header.append_paragraph('My header.')
        self.assertTrue(header.is_header)
        self.assertTrue(para.is_end_of_header_footer)
        # Create a footer and append a paragraph to it. The text in that paragraph
        # will appear at the bottom of every page of this section, below the main body text.
        footer = aw.HeaderFooter(doc, aw.HeaderFooterType.FOOTER_PRIMARY)
        doc.first_section.headers_footers.add(footer)
        para = footer.append_paragraph('My footer.')
        self.assertFalse(footer.is_header)
        self.assertTrue(para.is_end_of_header_footer)
        self.assertEqual(footer, para.parent_story)
        self.assertEqual(footer.parent_section, para.parent_section)
        self.assertEqual(footer.parent_section, header.parent_section)
        doc.save(file_name=ARTIFACTS_DIR + 'HeaderFooter.Create.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'HeaderFooter.Create.docx')
        self.assertTrue('My header.' in doc.first_section.headers_footers.get_by_header_footer_type(aw.HeaderFooterType.HEADER_PRIMARY).range.text)
        self.assertTrue('My footer.' in doc.first_section.headers_footers.get_by_header_footer_type(aw.HeaderFooterType.FOOTER_PRIMARY).range.text)

    def test_link(self):
        #ExStart
        #ExFor:HeaderFooter.is_linked_to_previous
        #ExFor:HeaderFooterCollection.__getitem__(int)
        #ExFor:HeaderFooterCollection.link_to_previous(HeaderFooterType,bool)
        #ExFor:HeaderFooterCollection.link_to_previous(bool)
        #ExFor:HeaderFooter.parent_section
        #ExSummary:Shows how to link headers and footers between sections.
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc=doc)
        builder.write('Section 1')
        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
        builder.write('Section 2')
        builder.insert_break(aw.BreakType.SECTION_BREAK_NEW_PAGE)
        builder.write('Section 3')
        # Move to the first section and create a header and a footer. By default,
        # the header and the footer will only appear on pages in the section that contains them.
        builder.move_to_section(0)
        builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY)
        builder.write('This is the header, which will be displayed in sections 1 and 2.')
        builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)
        builder.write('This is the footer, which will be displayed in sections 1, 2 and 3.')
        # We can link a section's headers/footers to the previous section's headers/footers
        # to allow the linking section to display the linked section's headers/footers.
        doc.sections[1].headers_footers.link_to_previous(is_link_to_previous=True)
        # Each section will still have its own header/footer objects. When we link sections,
        # the linking section will display the linked section's header/footers while keeping its own.
        self.assertNotEqual(doc.sections[0].headers_footers[0], doc.sections[1].headers_footers[0])
        self.assertNotEqual(doc.sections[0].headers_footers[0].parent_section, doc.sections[1].headers_footers[0].parent_section)
        # Link the headers/footers of the third section to the headers/footers of the second section.
        # The second section already links to the first section's header/footers,
        # so linking to the second section will create a link chain.
        # The first, second, and now the third sections will all display the first section's headers.
        doc.sections[2].headers_footers.link_to_previous(is_link_to_previous=True)
        # We can un-link a previous section's header/footers by passing "false" when calling the LinkToPrevious method.
        doc.sections[2].headers_footers.link_to_previous(is_link_to_previous=False)
        # We can also select only a specific type of header/footer to link using this method.
        # The third section now will have the same footer as the second and first sections, but not the header.
        doc.sections[2].headers_footers.link_to_previous(header_footer_type=aw.HeaderFooterType.FOOTER_PRIMARY, is_link_to_previous=True)
        # The first section's header/footers cannot link themselves to anything because there is no previous section.
        self.assertEqual(2, doc.sections[0].headers_footers.count)
        self.assertEqual(2, len(list(filter(lambda hf: not hf.as_header_footer().is_linked_to_previous, doc.sections[0].headers_footers))))
        # All the second section's header/footers are linked to the first section's headers/footers.
        self.assertEqual(6, doc.sections[1].headers_footers.count)
        self.assertEqual(6, len(list(filter(lambda hf: hf.as_header_footer().is_linked_to_previous, doc.sections[1].headers_footers))))
        # In the third section, only the footer is linked to the first section's footer via the second section.
        self.assertEqual(6, doc.sections[2].headers_footers.count)
        self.assertEqual(5, len(list(filter(lambda hf: not hf.as_header_footer().is_linked_to_previous, doc.sections[2].headers_footers))))
        self.assertTrue(doc.sections[2].headers_footers[3].is_linked_to_previous)
        doc.save(file_name=ARTIFACTS_DIR + 'HeaderFooter.Link.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'HeaderFooter.Link.docx')
        self.assertEqual(2, doc.sections[0].headers_footers.count)
        self.assertEqual(2, len(list(filter(lambda hf: not hf.as_header_footer().is_linked_to_previous, doc.sections[0].headers_footers))))
        self.assertEqual(0, doc.sections[1].headers_footers.count)
        self.assertEqual(0, len(list(filter(lambda hf: hf.as_header_footer().is_linked_to_previous, doc.sections[1].headers_footers))))
        self.assertEqual(5, doc.sections[2].headers_footers.count)
        self.assertEqual(5, len(list(filter(lambda hf: not hf.as_header_footer().is_linked_to_previous, doc.sections[2].headers_footers))))

    def test_export_mode(self):
        #ExStart
        #ExFor:HtmlSaveOptions.export_headers_footers_mode
        #ExFor:ExportHeadersFootersMode
        #ExSummary:Shows how to omit headers/footers when saving a document to HTML.
        doc = aw.Document(file_name=MY_DIR + 'Header and footer types.docx')
        # This document contains headers and footers. We can access them via the "HeadersFooters" collection.
        self.assertEqual('First header', doc.first_section.headers_footers.get_by_header_footer_type(aw.HeaderFooterType.HEADER_FIRST).get_text().strip())
        # Formats such as .html do not split the document into pages, so headers/footers will not function the same way
        # they would when we open the document as a .docx using Microsoft Word.
        # If we convert a document with headers/footers to html, the conversion will assimilate the headers/footers into body text.
        # We can use a SaveOptions object to omit headers/footers while converting to html.
        save_options = aw.saving.HtmlSaveOptions(aw.SaveFormat.HTML)
        save_options.export_headers_footers_mode = aw.saving.ExportHeadersFootersMode.NONE
        doc.save(file_name=ARTIFACTS_DIR + 'HeaderFooter.ExportMode.html', save_options=save_options)
        # Open our saved document and verify that it does not contain the header's text
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'HeaderFooter.ExportMode.html')
        self.assertFalse('First header' in doc.range.text)
        #ExEnd

    def test_replace_text(self):
        #ExStart
        #ExFor:Document.first_section
        #ExFor:Section.headers_footers
        #ExFor:HeaderFooterCollection.__getitem__(HeaderFooterType)
        #ExFor:HeaderFooter
        #ExFor:Range.replace(str,str,FindReplaceOptions)
        #ExSummary:Shows how to replace text in a document's footer.
        doc = aw.Document(file_name=MY_DIR + 'Footer.docx')
        headers_footers = doc.first_section.headers_footers
        footer = headers_footers.get_by_header_footer_type(aw.HeaderFooterType.FOOTER_PRIMARY)
        options = aw.replacing.FindReplaceOptions()
        options.match_case = False
        options.find_whole_words_only = False
        current_year = datetime.datetime.now().year
        footer.range.replace(pattern='(C) 2006 Aspose Pty Ltd.', replacement=f'Copyright (C) {current_year} by Aspose Pty Ltd.', options=options)
        doc.save(file_name=ARTIFACTS_DIR + 'HeaderFooter.ReplaceText.docx')
        #ExEnd
        doc = aw.Document(file_name=ARTIFACTS_DIR + 'HeaderFooter.ReplaceText.docx')
        self.assertTrue(f'Copyright (C) {current_year} by Aspose Pty Ltd.' in doc.range.text)

    def test_remove_footers(self):
        #ExStart
        #ExFor:Section.headers_footers
        #ExFor:HeaderFooterCollection
        #ExFor:HeaderFooterCollection.__getitem__(HeaderFooterType)
        #ExFor:HeaderFooter
        #ExSummary:Shows how to delete all footers from a document.
        doc = aw.Document(MY_DIR + 'Header and footer types.docx')
        # Iterate through each section and remove footers of every kind.
        for section in doc:
            section = section.as_section()
            # There are three kinds of footer and header types.
            # 1 -  The "First" header/footer, which only appears on the first page of a section.
            footer = section.headers_footers[aw.HeaderFooterType.FOOTER_FIRST]
            if footer is not None:
                footer.remove()
            # 2 -  The "Primary" header/footer, which appears on odd pages.
            footer = section.headers_footers[aw.HeaderFooterType.FOOTER_PRIMARY]
            if footer is not None:
                footer.remove()
            # 3 -  The "Even" header/footer, which appears on even pages.
            footer = section.headers_footers[aw.HeaderFooterType.FOOTER_EVEN]
            if footer is not None:
                footer.remove()
            self.assertEqual(0, len([node for node in section.headers_footers if not node.as_header_footer().is_header]))
        doc.save(ARTIFACTS_DIR + 'HeaderFooter.remove_footers.docx')
        #ExEnd
        doc = aw.Document(ARTIFACTS_DIR + 'HeaderFooter.remove_footers.docx')
        self.assertEqual(1, doc.sections.count)
        self.assertEqual(0, len([node for node in doc.first_section.headers_footers if not node.as_header_footer().is_header]))
        self.assertEqual(3, len([node for node in doc.first_section.headers_footers if node.as_header_footer().is_header]))