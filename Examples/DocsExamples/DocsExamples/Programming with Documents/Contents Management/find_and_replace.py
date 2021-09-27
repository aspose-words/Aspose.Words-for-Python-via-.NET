import unittest
import os
import sys
from datetime import date, datetime

base_dir = os.path.abspath(os.curdir) + "/"
base_dir = base_dir[:base_dir.find("Aspose.Words-for-Python-via-.NET")]
base_dir = base_dir + "Aspose.Words-for-Python-via-.NET/Examples/DocsExamples/DocsExamples"
sys.path.insert(0, base_dir)

import docs_examples_base as docs_base

import aspose.words as aw
import aspose.pydrawing as drawing

class FindAndReplace(docs_base.DocsExamplesBase):
    
    def test_simple_find_replace(self) :
        
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.writeln("Hello _CustomerName_,")
        print("Original document text: " + doc.range.text)

        doc.range.replace("_CustomerName_", "James Bond", aw.replacing.FindReplaceOptions(aw.replacing.FindReplaceDirection.FORWARD))

        print("Document text after replace: " + doc.range.text)

        # Save the modified document
        doc.save(docs_base.artifacts_dir + "FindAndReplace.simple_find_replace.docx")
        

    def test_find_and_highlight(self) :
        
        #ExStart:FindAndHighlight
        doc = aw.Document(docs_base.my_dir + "Find and highlight.docx")

        options = aw.replacing.FindReplaceOptions()
        options.direction = aw.replacing.FindReplaceDirection.BACKWARD
        options.apply_font.highlight_color = drawing.Color.yellow

        text = "your document"
        doc.range.replace(text, text, options)

        doc.save(docs_base.artifacts_dir + "FindAndReplace.find_and_highlight.docx")
        #ExEnd:FindAndHighlight
        

    def test_meta_characters_in_search_pattern(self) :
        
        # meta-characters
        # &p - paragraph break
        # &b - section break
        # &m - page break
        # &l - manual line break
        

        #ExStart:MetaCharactersInSearchPattern
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
            
        builder.writeln("This is Line 1")
        builder.writeln("This is Line 2")

        doc.range.replace("This is Line 1&pThis is Line 2", "This is replaced line")

        builder.move_to_document_end()
        builder.write("This is Line 1")
        builder.insert_break(aw.BreakType.PAGE_BREAK)
        builder.writeln("This is Line 2")

        doc.range.replace("This is Line 1&mThis is Line 2", "Page break is replaced with new text.")

        doc.save(docs_base.artifacts_dir + "FindAndReplace.meta_characters_in_search_pattern.docx")
        #ExEnd:MetaCharactersInSearchPattern
        

    def test_replace_text_containing_meta_characters(self) :
        
        #ExStart:ReplaceTextContainingMetaCharacters
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.font.name = "Arial"
        builder.writeln("First section")
        builder.writeln("  1st paragraph")
        builder.writeln("  2nd paragraph")
        builder.writeln("insert-section")
        builder.writeln("Second section")
        builder.writeln("  1st paragraph")

        findReplaceOptions = aw.replacing.FindReplaceOptions()
        findReplaceOptions.apply_paragraph_format.alignment = aw.ParagraphAlignment.CENTER

        # Double each paragraph break after word "section", add kind of underline and make it centered.
        count = doc.range.replace("section&p", "section&p----------------------&p", findReplaceOptions)

        # Insert section break instead of custom text tag.
        count = doc.range.replace("insert-section", "&b", findReplaceOptions)

        doc.save(docs_base.artifacts_dir + "FindAndReplace.replace_text_containing_meta_characters.docx")
        #ExEnd:ReplaceTextContainingMetaCharacters
        

    def test_ignore_text_inside_fields(self) :
        
        #ExStart:IgnoreTextInsideFields
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert field with text inside.
        builder.insert_field("INCLUDETEXT", "Text in field")
            
        options = aw.replacing.FindReplaceOptions()
        options.ignore_fields = True 
            
        doc.range.replace_regex("e", "*", options)
            
        print(doc.get_text())

        options.ignore_fields = False
        doc.range.replace("e", "*", options)
            
        print(doc.get_text())
        #ExEnd:IgnoreTextInsideFields
        

    @unittest.skip("Regular expressions is not supported yet.")
    def test_ignore_text_inside_delete_revisions(self) :
        
        #ExStart:IgnoreTextInsideDeleteRevisions
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert non-revised text.
        builder.writeln("Deleted")
        builder.write("Text")

        # Remove first paragraph with tracking revisions.
        doc.start_track_revisions("author", DateTime.now)
        doc.first_section.body.first_paragraph.remove()
        doc.stop_track_revisions()

        options = aw.replacing.FindReplaceOptions()
        options.ignore_deleted = True 

        doc.range.replace_regex("e", "*", options)

        print(doc.get_text())

        options.ignore_deleted = False
        doc.range.replace(regex, "*", options)

        print(doc.get_text())
        #ExEnd:IgnoreTextInsideDeleteRevisions
        

    def test_ignore_text_inside_insert_revisions(self) :
        
        #ExStart:IgnoreTextInsideInsertRevisions
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        # Insert text with tracking revisions.
        doc.start_track_revisions("author", datetime.today())
        builder.writeln("Inserted")
        doc.stop_track_revisions()

        # Insert non-revised text.
        builder.write("Text")

        options = aw.replacing.FindReplaceOptions()
        options.ignore_inserted = True 

        doc.range.replace_regex("e", "*", options)
            
        print(doc.get_text())

        options.ignore_inserted = False
        doc.range.replace("e", "*", options)
            
        print(doc.get_text())
        #ExEnd:IgnoreTextInsideInsertRevisions
        

    def test_replace_text_in_footer(self) :
        
        #ExStart:ReplaceTextInFooter
        doc = aw.Document(docs_base.my_dir + "Footer.docx")

        headersFooters = doc.first_section.headers_footers
        footer = headersFooters.get_by_header_footer_type(aw.HeaderFooterType.FOOTER_PRIMARY)

        options = aw.replacing.FindReplaceOptions()
        options.match_case = False
        options.find_whole_words_only = False 

        footer.range.replace("(C) 2006 Aspose Pty Ltd.", "Copyright (C) 2020 by Aspose Pty Ltd.", options)

        doc.save(docs_base.artifacts_dir + "FindAndReplace.replace_text_in_footer.docx")
        #ExEnd:ReplaceTextInFooter
        

    @unittest.skip("Regular expressions is not supported yet.")
    def test_replace_with_regex(self) :
        
        #ExStart:ReplaceWithRegex
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
            
        builder.writeln("sad mad bad")

        options = aw.replacing.FindReplaceOptions()

        doc.range.replace_regex("[s|m]ad", "bad", options)

        doc.save(docs_base.artifacts_dir + "FindAndReplace.replace_with_regex.docx")
        #ExEnd:ReplaceWithRegex
        
    
    def test_recognize_and_substitutions_within_replacement_patterns(self) :
        
        #ExStart:RecognizeAndSubstitutionsWithinReplacementPatterns
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)

        builder.write("Jason give money to Paul.")

        options = aw.replacing.FindReplaceOptions()
        options.use_substitutions = True 

        doc.range.replace_regex("([A-z]+) give money to ([A-z]+)", "$2 take money from $1", options)
        #ExEnd:RecognizeAndSubstitutionsWithinReplacementPatterns
        

    def test_replace_with_string(self) :
        
        #ExStart:ReplaceWithString
        doc = aw.Document()
        builder = aw.DocumentBuilder(doc)
            
        builder.writeln("sad mad bad")

        doc.range.replace("sad", "bad", aw.replacing.FindReplaceOptions(aw.replacing.FindReplaceDirection.FORWARD))

        doc.save(docs_base.artifacts_dir + "FindAndReplace.replace_with_string.docx")
        #ExEnd:ReplaceWithString
        

    def test_replace_text_in_table(self) :
        
        #ExStart:ReplaceText
        doc = aw.Document(docs_base.my_dir + "Tables.docx")

        table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()

        table.range.replace("Carrots", "Eggs", aw.replacing.FindReplaceOptions(aw.replacing.FindReplaceDirection.FORWARD))
        table.last_row.last_cell.range.replace("50", "20", aw.replacing.FindReplaceOptions(aw.replacing.FindReplaceDirection.FORWARD))

        doc.save(docs_base.artifacts_dir + "FindAndReplace.replace_text_in_table.docx")
        #ExEnd:ReplaceText
        
    


if __name__ == '__main__':
    unittest.main()