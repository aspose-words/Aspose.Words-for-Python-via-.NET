import unittest

import api_example_base as aeb
from document_helper import DocumentHelper

import aspose.words as aw

class ExCompatibilityOptions(aeb.ApiExampleBase):
    
    #ExStart
    #ExFor:Compatibility
    #ExFor:CompatibilityOptions
    #ExFor:CompatibilityOptions.optimize_for(MsWordVersion)
    #ExFor:Document.compatibility_options
    #ExFor:MsWordVersion
    #ExSummary:Shows how to optimize the document for different versions of Microsoft Word.
    def test_optimize_for(self) :
        
        doc = aw.Document()

        # This object contains an extensive list of flags unique to each document
        # that allow us to facilitate backward compatibility with older versions of Microsoft Word.
        options = doc.compatibility_options

        # Print the default settings for a blank document.
        print("\nDefault optimization settings:")
        self.print_compatibility_options(options)

        # We can access these settings in Microsoft Word via "File" -> "Options" -> "Advanced" -> "Compatibility options for...".
        doc.save(aeb.artifacts_dir + "CompatibilityOptions.optimize_for.default_settings.docx")

        # We can use the OptimizeFor method to ensure optimal compatibility with a specific Microsoft Word version.
        doc.compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2010)
        print("\nOptimized for Word 2010:")
        self.print_compatibility_options(options)

        doc.compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2000)
        print("\nOptimized for Word 2000:")
        self.print_compatibility_options(options)
        

    # <summary>
    # Groups all flags in a document's compatibility options object by state, then prints each group.
    # </summary>
    @staticmethod
    def print_compatibility_options(options) : 
        
        for i in [1,0] :
            print("\tEnabled options:" if i==1 else "\tDisabled options:")
            test = i==1
            for opt in dir(options):
               if not opt.startswith('__') and not callable(getattr(options, opt)) and getattr(options, opt)==test :
                    print(f"\t\t{opt}")
                
            
        
    #ExEnd

    def test_tables(self) :
        
        doc = aw.Document()

        compatibilityOptions = doc.compatibility_options
        compatibilityOptions.optimize_for(aw.settings.MsWordVersion.WORD2002)

        self.assertEqual(False, compatibilityOptions.adjust_line_height_in_table)
        self.assertEqual(False, compatibilityOptions.align_tables_row_by_row)
        self.assertEqual(True, compatibilityOptions.allow_space_of_same_style_in_table)
        self.assertEqual(True, compatibilityOptions.do_not_autofit_constrained_tables)
        self.assertEqual(True, compatibilityOptions.do_not_break_constrained_forced_table)
        self.assertEqual(False, compatibilityOptions.do_not_break_wrapped_tables)
        self.assertEqual(False, compatibilityOptions.do_not_snap_to_grid_in_cell)
        self.assertEqual(False, compatibilityOptions.do_not_use_htmlparagraph_auto_spacing)
        self.assertEqual(True, compatibilityOptions.do_not_vert_align_cell_with_sp)
        self.assertEqual(False, compatibilityOptions.forget_last_tab_alignment)
        self.assertEqual(True, compatibilityOptions.grow_autofit)
        self.assertEqual(False, compatibilityOptions.layout_raw_table_width)
        self.assertEqual(False, compatibilityOptions.layout_table_rows_apart)
        self.assertEqual(False, compatibilityOptions.no_column_balance)
        self.assertEqual(False, compatibilityOptions.override_table_style_font_size_and_justification)
        self.assertEqual(False, compatibilityOptions.use_single_borderfor_contiguous_cells)
        self.assertEqual(True, compatibilityOptions.use_word2002_table_style_rules)
        self.assertEqual(False, compatibilityOptions.use_word2010_table_style_rules)

        # In the output document, these settings can be accessed in Microsoft Word via
        # File -> Options -> Advanced -> Compatibility options for...
        doc.save(aeb.artifacts_dir + "CompatibilityOptions.tables.docx")
        

    def test_breaks(self) :
        
        doc = aw.Document()

        compatibilityOptions = doc.compatibility_options
        compatibilityOptions.optimize_for(aw.settings.MsWordVersion.WORD2000)

        self.assertEqual(False, compatibilityOptions.apply_breaking_rules)
        self.assertEqual(True, compatibilityOptions.do_not_use_east_asian_break_rules)
        self.assertEqual(False, compatibilityOptions.show_breaks_in_frames)
        self.assertEqual(True, compatibilityOptions.split_pg_break_and_para_mark)
        self.assertEqual(True, compatibilityOptions.use_alt_kinsoku_line_break_rules)
        self.assertEqual(False, compatibilityOptions.use_word97_line_break_rules)

        # In the output document, these settings can be accessed in Microsoft Word via
        # File -> Options -> Advanced -> Compatibility options for...
        doc.save(aeb.artifacts_dir + "CompatibilityOptions.breaks.docx")
        

    def test_spacing(self) :
        
        doc = aw.Document()

        compatibilityOptions = doc.compatibility_options
        compatibilityOptions.optimize_for(aw.settings.MsWordVersion.WORD2000)

        self.assertEqual(False, compatibilityOptions.auto_space_like_word95)
        self.assertEqual(True, compatibilityOptions.display_hangul_fixed_width)
        self.assertEqual(False, compatibilityOptions.no_extra_line_spacing)
        self.assertEqual(False, compatibilityOptions.no_leading)
        self.assertEqual(False, compatibilityOptions.no_space_raise_lower)
        self.assertEqual(False, compatibilityOptions.space_for_ul)
        self.assertEqual(False, compatibilityOptions.spacing_in_whole_points)
        self.assertEqual(False, compatibilityOptions.suppress_bottom_spacing)
        self.assertEqual(False, compatibilityOptions.suppress_sp_bf_after_pg_brk)
        self.assertEqual(False, compatibilityOptions.suppress_spacing_at_top_of_page)
        self.assertEqual(False, compatibilityOptions.suppress_top_spacing)
        self.assertEqual(False, compatibilityOptions.ul_trail_space)

        # In the output document, these settings can be accessed in Microsoft Word via
        # File -> Options -> Advanced -> Compatibility options for...
        doc.save(aeb.artifacts_dir + "CompatibilityOptions.spacing.docx")
        

    def test_word_perfect(self) :
        
        doc = aw.Document()

        compatibilityOptions = doc.compatibility_options
        compatibilityOptions.optimize_for(aw.settings.MsWordVersion.WORD2000)

        self.assertEqual(False, compatibilityOptions.suppress_top_spacing_wp)
        self.assertEqual(False, compatibilityOptions.truncate_font_heights_like_wp6)
        self.assertEqual(False, compatibilityOptions.wpjustification)
        self.assertEqual(False, compatibilityOptions.wpspace_width)
        self.assertEqual(False, compatibilityOptions.wrap_trail_spaces)

        # In the output document, these settings can be accessed in Microsoft Word via
        # File -> Options -> Advanced -> Compatibility options for...
        doc.save(aeb.artifacts_dir + "CompatibilityOptions.word_perfect.docx")
        

    def test_alignment(self) :
        
        doc = aw.Document()
            
        compatibilityOptions = doc.compatibility_options
        compatibilityOptions.optimize_for(aw.settings.MsWordVersion.WORD2000)

        self.assertEqual(True, compatibilityOptions.cached_col_balance)
        self.assertEqual(True, compatibilityOptions.do_not_vert_align_in_txbx)
        self.assertEqual(True, compatibilityOptions.do_not_wrap_text_with_punct)
        self.assertEqual(False, compatibilityOptions.no_tab_hang_ind)

        # In the output document, these settings can be accessed in Microsoft Word via
        # File -> Options -> Advanced -> Compatibility options for...
        doc.save(aeb.artifacts_dir + "CompatibilityOptions.alignment.docx")
        

    def test_legacy(self) :
        
        doc = aw.Document()

        compatibilityOptions = doc.compatibility_options
        compatibilityOptions.optimize_for(aw.settings.MsWordVersion.WORD2000)

        self.assertEqual(False, compatibilityOptions.footnote_layout_like_ww8)
        self.assertEqual(False, compatibilityOptions.line_wrap_like_word6)
        self.assertEqual(False, compatibilityOptions.mwsmall_caps)
        self.assertEqual(False, compatibilityOptions.shape_layout_like_ww8)
        self.assertEqual(False, compatibilityOptions.uicompat97_to2003)

        # In the output document, these settings can be accessed in Microsoft Word via
        # File -> Options -> Advanced -> Compatibility options for...
        doc.save(aeb.artifacts_dir + "CompatibilityOptions.legacy.docx")
        

    def test_list(self) :
        
        doc = aw.Document()

        compatibilityOptions = doc.compatibility_options
        compatibilityOptions.optimize_for(aw.settings.MsWordVersion.WORD2000)

        self.assertEqual(True, compatibilityOptions.underline_tab_in_num_list)
        self.assertEqual(True, compatibilityOptions.use_normal_style_for_list)

        # In the output document, these settings can be accessed in Microsoft Word via
        # File -> Options -> Advanced -> Compatibility options for...
        doc.save(aeb.artifacts_dir + "CompatibilityOptions.list.docx")
        

    def test_misc(self) :
        
        doc = aw.Document()

        compatibilityOptions = doc.compatibility_options
        compatibilityOptions.optimize_for(aw.settings.MsWordVersion.WORD2000)

        self.assertEqual(False, compatibilityOptions.balance_single_byte_double_byte_width)
        self.assertEqual(False, compatibilityOptions.conv_mail_merge_esc)
        self.assertEqual(False, compatibilityOptions.do_not_expand_shift_return)
        self.assertEqual(False, compatibilityOptions.do_not_leave_backslash_alone)
        self.assertEqual(False, compatibilityOptions.do_not_suppress_paragraph_borders)
        self.assertEqual(True, compatibilityOptions.do_not_use_indent_as_numbering_tab_stop)
        self.assertEqual(False, compatibilityOptions.print_body_text_before_header)
        self.assertEqual(False, compatibilityOptions.print_col_black)
        self.assertEqual(True, compatibilityOptions.select_fld_with_first_or_last_char)
        self.assertEqual(False, compatibilityOptions.sub_font_by_size)
        self.assertEqual(False, compatibilityOptions.swap_borders_facing_pgs)
        self.assertEqual(False, compatibilityOptions.transparent_metafiles)
        self.assertEqual(True, compatibilityOptions.use_ansi_kerning_pairs)
        self.assertEqual(False, compatibilityOptions.use_felayout)
        self.assertEqual(False, compatibilityOptions.use_printer_metrics)

        # In the output document, these settings can be accessed in Microsoft Word via
        # File -> Options -> Advanced -> Compatibility options for...
        doc.save(aeb.artifacts_dir + "CompatibilityOptions.misc.docx")
        
    
if __name__ == '__main__':
    unittest.main()    
