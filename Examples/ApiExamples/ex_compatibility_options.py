# -*- coding: utf-8 -*-
# Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#
# This file is part of Aspose.Words. The source code in this file
# is only intended as a supplement to the documentation, and is provided
# "as is", without warranty of any kind, either expressed or implied.
#####################################
import aspose.words as aw
import aspose.words.settings
import unittest
from api_example_base import ApiExampleBase, ARTIFACTS_DIR, MY_DIR

class ExCompatibilityOptions(ApiExampleBase):

    def test_tables(self):
        doc = aw.Document()
        compatibility_options = doc.compatibility_options
        compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2002)
        self.assertEqual(False, compatibility_options.adjust_line_height_in_table)
        self.assertEqual(False, compatibility_options.align_tables_row_by_row)
        self.assertEqual(True, compatibility_options.allow_space_of_same_style_in_table)
        self.assertEqual(True, compatibility_options.do_not_autofit_constrained_tables)
        self.assertEqual(True, compatibility_options.do_not_break_constrained_forced_table)
        self.assertEqual(False, compatibility_options.do_not_break_wrapped_tables)
        self.assertEqual(False, compatibility_options.do_not_snap_to_grid_in_cell)
        self.assertEqual(False, compatibility_options.do_not_use_html_paragraph_auto_spacing)
        self.assertEqual(True, compatibility_options.do_not_vert_align_cell_with_sp)
        self.assertEqual(False, compatibility_options.forget_last_tab_alignment)
        self.assertEqual(True, compatibility_options.grow_autofit)
        self.assertEqual(False, compatibility_options.layout_raw_table_width)
        self.assertEqual(False, compatibility_options.layout_table_rows_apart)
        self.assertEqual(False, compatibility_options.no_column_balance)
        self.assertEqual(False, compatibility_options.override_table_style_font_size_and_justification)
        self.assertEqual(False, compatibility_options.use_single_borderfor_contiguous_cells)
        self.assertEqual(True, compatibility_options.use_word2002_table_style_rules)
        self.assertEqual(False, compatibility_options.use_word2010_table_style_rules)
        # In the output document, these settings can be accessed in Microsoft Word via
        # File -> Options -> Advanced -> Compatibility options for...
        doc.save(file_name=ARTIFACTS_DIR + 'CompatibilityOptions.Tables.docx')

    def test_breaks(self):
        doc = aw.Document()
        compatibility_options = doc.compatibility_options
        compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2000)
        self.assertEqual(False, compatibility_options.apply_breaking_rules)
        self.assertEqual(True, compatibility_options.do_not_use_east_asian_break_rules)
        self.assertEqual(False, compatibility_options.show_breaks_in_frames)
        self.assertEqual(True, compatibility_options.split_pg_break_and_para_mark)
        self.assertEqual(True, compatibility_options.use_alt_kinsoku_line_break_rules)
        self.assertEqual(False, compatibility_options.use_word97_line_break_rules)
        # In the output document, these settings can be accessed in Microsoft Word via
        # File -> Options -> Advanced -> Compatibility options for...
        doc.save(file_name=ARTIFACTS_DIR + 'CompatibilityOptions.Breaks.docx')

    def test_spacing(self):
        doc = aw.Document()
        compatibility_options = doc.compatibility_options
        compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2000)
        self.assertEqual(False, compatibility_options.auto_space_like_word95)
        self.assertEqual(True, compatibility_options.display_hangul_fixed_width)
        self.assertEqual(False, compatibility_options.no_extra_line_spacing)
        self.assertEqual(False, compatibility_options.no_leading)
        self.assertEqual(False, compatibility_options.no_space_raise_lower)
        self.assertEqual(False, compatibility_options.space_for_ul)
        self.assertEqual(False, compatibility_options.spacing_in_whole_points)
        self.assertEqual(False, compatibility_options.suppress_bottom_spacing)
        self.assertEqual(False, compatibility_options.suppress_sp_bf_after_pg_brk)
        self.assertEqual(False, compatibility_options.suppress_spacing_at_top_of_page)
        self.assertEqual(False, compatibility_options.suppress_top_spacing)
        self.assertEqual(False, compatibility_options.ul_trail_space)
        # In the output document, these settings can be accessed in Microsoft Word via
        # File -> Options -> Advanced -> Compatibility options for...
        doc.save(file_name=ARTIFACTS_DIR + 'CompatibilityOptions.Spacing.docx')

    def test_word_perfect(self):
        doc = aw.Document()
        compatibility_options = doc.compatibility_options
        compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2000)
        self.assertEqual(False, compatibility_options.suppress_top_spacing_wp)
        self.assertEqual(False, compatibility_options.truncate_font_heights_like_wp6)
        self.assertEqual(False, compatibility_options.wp_justification)
        self.assertEqual(False, compatibility_options.wp_space_width)
        self.assertEqual(False, compatibility_options.wrap_trail_spaces)
        # In the output document, these settings can be accessed in Microsoft Word via
        # File -> Options -> Advanced -> Compatibility options for...
        doc.save(file_name=ARTIFACTS_DIR + 'CompatibilityOptions.WordPerfect.docx')

    def test_alignment(self):
        doc = aw.Document()
        compatibility_options = doc.compatibility_options
        compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2000)
        self.assertEqual(True, compatibility_options.cached_col_balance)
        self.assertEqual(True, compatibility_options.do_not_vert_align_in_txbx)
        self.assertEqual(True, compatibility_options.do_not_wrap_text_with_punct)
        self.assertEqual(False, compatibility_options.no_tab_hang_ind)
        # In the output document, these settings can be accessed in Microsoft Word via
        # File -> Options -> Advanced -> Compatibility options for...
        doc.save(file_name=ARTIFACTS_DIR + 'CompatibilityOptions.Alignment.docx')

    def test_legacy(self):
        doc = aw.Document()
        compatibility_options = doc.compatibility_options
        compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2000)
        self.assertEqual(False, compatibility_options.footnote_layout_like_ww8)
        self.assertEqual(False, compatibility_options.line_wrap_like_word6)
        self.assertEqual(False, compatibility_options.mw_small_caps)
        self.assertEqual(False, compatibility_options.shape_layout_like_ww8)
        self.assertEqual(False, compatibility_options.ui_compat_97_to_2003)
        # In the output document, these settings can be accessed in Microsoft Word via
        # File -> Options -> Advanced -> Compatibility options for...
        doc.save(file_name=ARTIFACTS_DIR + 'CompatibilityOptions.Legacy.docx')

    def test_list(self):
        doc = aw.Document()
        compatibility_options = doc.compatibility_options
        compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2000)
        self.assertEqual(True, compatibility_options.underline_tab_in_num_list)
        self.assertEqual(True, compatibility_options.use_normal_style_for_list)
        # In the output document, these settings can be accessed in Microsoft Word via
        # File -> Options -> Advanced -> Compatibility options for...
        doc.save(file_name=ARTIFACTS_DIR + 'CompatibilityOptions.List.docx')

    def test_misc(self):
        doc = aw.Document()
        compatibility_options = doc.compatibility_options
        compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2000)
        self.assertEqual(False, compatibility_options.balance_single_byte_double_byte_width)
        self.assertEqual(False, compatibility_options.conv_mail_merge_esc)
        self.assertEqual(False, compatibility_options.do_not_expand_shift_return)
        self.assertEqual(False, compatibility_options.do_not_leave_backslash_alone)
        self.assertEqual(False, compatibility_options.do_not_suppress_paragraph_borders)
        self.assertEqual(True, compatibility_options.do_not_use_indent_as_numbering_tab_stop)
        self.assertEqual(False, compatibility_options.print_body_text_before_header)
        self.assertEqual(False, compatibility_options.print_col_black)
        self.assertEqual(True, compatibility_options.select_fld_with_first_or_last_char)
        self.assertEqual(False, compatibility_options.sub_font_by_size)
        self.assertEqual(False, compatibility_options.swap_borders_facing_pgs)
        self.assertEqual(False, compatibility_options.transparent_metafiles)
        self.assertEqual(True, compatibility_options.use_ansi_kerning_pairs)
        self.assertEqual(False, compatibility_options.use_fe_layout)
        self.assertEqual(False, compatibility_options.use_printer_metrics)
        # In the output document, these settings can be accessed in Microsoft Word via
        # File -> Options -> Advanced -> Compatibility options for...
        doc.save(file_name=ARTIFACTS_DIR + 'CompatibilityOptions.Misc.docx')

    def test_optimize_for(self):
        #ExStart
        #ExFor:Compatibility
        #ExFor:CompatibilityOptions
        #ExFor:CompatibilityOptions.optimize_for(MsWordVersion)
        #ExFor:Document.compatibility_options
        #ExFor:MsWordVersion
        #ExSummary:Shows how to optimize the document for different versions of Microsoft Word.

        def optimize_for():
            doc = aw.Document()
            # This object contains an extensive list of flags unique to each document
            # that allow us to facilitate backward compatibility with older versions of Microsoft Word.
            options = doc.compatibility_options
            # Print the default settings for a blank document.
            print('\nDefault optimization settings:')
            print_compatibility_options(options)
            # We can access these settings in Microsoft Word via "File" -> "Options" -> "Advanced" -> "Compatibility options for...".
            doc.save(ARTIFACTS_DIR + 'CompatibilityOptions.optimize_for.default_settings.docx')
            # We can use the OptimizeFor method to ensure optimal compatibility with a specific Microsoft Word version.
            doc.compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2010)
            print('\nOptimized for Word 2010:')
            print_compatibility_options(options)
            doc.compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2000)
            print('\nOptimized for Word 2000:')
            print_compatibility_options(options)

        def print_compatibility_options(options: aw.settings.CompatibilityOptions):
            """Groups all flags in a document's compatibility options object by state, then prints each group."""
            for enabled in (True, False):
                print('\tEnabled options:' if enabled else '\tDisabled options:')
                for opt in dir(options):
                    if not opt.startswith('__') and (not callable(getattr(options, opt))) and (getattr(options, opt) == enabled):
                        print(f'\t\t{opt}')
        #ExEnd
        optimize_for()