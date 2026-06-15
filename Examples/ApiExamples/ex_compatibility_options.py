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
    #ExStart
    #ExFor:Compatibility
    #ExFor:CompatibilityOptions
    #ExFor:CompatibilityOptions.optimize_for(MsWordVersion)
    #ExFor:Document.compatibility_options
    #ExFor:MsWordVersion
    #ExFor:CompatibilityOptions.adjust_line_height_in_table
    #ExFor:CompatibilityOptions.align_tables_row_by_row
    #ExFor:CompatibilityOptions.allow_space_of_same_style_in_table
    #ExFor:CompatibilityOptions.apply_breaking_rules
    #ExFor:CompatibilityOptions.autofit_to_first_fixed_width_cell
    #ExFor:CompatibilityOptions.auto_space_like_word95
    #ExFor:CompatibilityOptions.balance_single_byte_double_byte_width
    #ExFor:CompatibilityOptions.cached_col_balance
    #ExFor:CompatibilityOptions.conv_mail_merge_esc
    #ExFor:CompatibilityOptions.disable_open_type_font_formatting_features
    #ExFor:CompatibilityOptions.display_hangul_fixed_width
    #ExFor:CompatibilityOptions.do_not_autofit_constrained_tables
    #ExFor:CompatibilityOptions.do_not_break_constrained_forced_table
    #ExFor:CompatibilityOptions.do_not_break_wrapped_tables
    #ExFor:CompatibilityOptions.do_not_expand_shift_return
    #ExFor:CompatibilityOptions.do_not_leave_backslash_alone
    #ExFor:CompatibilityOptions.do_not_snap_to_grid_in_cell
    #ExFor:CompatibilityOptions.do_not_suppress_indentation
    #ExFor:CompatibilityOptions.do_not_suppress_paragraph_borders
    #ExFor:CompatibilityOptions.do_not_use_east_asian_break_rules
    #ExFor:CompatibilityOptions.do_not_use_html_paragraph_auto_spacing
    #ExFor:CompatibilityOptions.do_not_use_indent_as_numbering_tab_stop
    #ExFor:CompatibilityOptions.do_not_vert_align_cell_with_sp
    #ExFor:CompatibilityOptions.do_not_vert_align_in_txbx
    #ExFor:CompatibilityOptions.do_not_wrap_text_with_punct
    #ExFor:CompatibilityOptions.footnote_layout_like_ww8
    #ExFor:CompatibilityOptions.forget_last_tab_alignment
    #ExFor:CompatibilityOptions.grow_autofit
    #ExFor:CompatibilityOptions.layout_raw_table_width
    #ExFor:CompatibilityOptions.layout_table_rows_apart
    #ExFor:CompatibilityOptions.line_wrap_like_word6
    #ExFor:CompatibilityOptions.mw_small_caps
    #ExFor:CompatibilityOptions.no_column_balance
    #ExFor:CompatibilityOptions.no_extra_line_spacing
    #ExFor:CompatibilityOptions.no_leading
    #ExFor:CompatibilityOptions.no_space_raise_lower
    #ExFor:CompatibilityOptions.no_tab_hang_ind
    #ExFor:CompatibilityOptions.override_table_style_font_size_and_justification
    #ExFor:CompatibilityOptions.print_body_text_before_header
    #ExFor:CompatibilityOptions.print_col_black
    #ExFor:CompatibilityOptions.select_fld_with_first_or_last_char
    #ExFor:CompatibilityOptions.shape_layout_like_ww8
    #ExFor:CompatibilityOptions.show_breaks_in_frames
    #ExFor:CompatibilityOptions.space_for_ul
    #ExFor:CompatibilityOptions.spacing_in_whole_points
    #ExFor:CompatibilityOptions.split_pg_break_and_para_mark
    #ExFor:CompatibilityOptions.sub_font_by_size
    #ExFor:CompatibilityOptions.suppress_bottom_spacing
    #ExFor:CompatibilityOptions.suppress_spacing_at_top_of_page
    #ExFor:CompatibilityOptions.suppress_sp_bf_after_pg_brk
    #ExFor:CompatibilityOptions.suppress_top_spacing
    #ExFor:CompatibilityOptions.suppress_top_spacing_wp
    #ExFor:CompatibilityOptions.swap_borders_facing_pgs
    #ExFor:CompatibilityOptions.swap_inside_and_outside_for_mirror_indents_and_relative_positioning
    #ExFor:CompatibilityOptions.transparent_metafiles
    #ExFor:CompatibilityOptions.truncate_font_heights_like_wp6
    #ExFor:CompatibilityOptions.ui_compat_97_to_2003
    #ExFor:CompatibilityOptions.ul_trail_space
    #ExFor:CompatibilityOptions.underline_tab_in_num_list
    #ExFor:CompatibilityOptions.use_alt_kinsoku_line_break_rules
    #ExFor:CompatibilityOptions.use_ansi_kerning_pairs
    #ExFor:CompatibilityOptions.use_fe_layout
    #ExFor:CompatibilityOptions.use_normal_style_for_list
    #ExFor:CompatibilityOptions.use_printer_metrics
    #ExFor:CompatibilityOptions.use_single_borderfor_contiguous_cells
    #ExFor:CompatibilityOptions.use_word2002_table_style_rules
    #ExFor:CompatibilityOptions.use_word2010_table_style_rules
    #ExFor:CompatibilityOptions.use_word97_line_break_rules
    #ExFor:CompatibilityOptions.wp_justification
    #ExFor:CompatibilityOptions.wp_space_width
    #ExFor:CompatibilityOptions.wrap_trail_spaces
    #ExSummary:Shows how to optimize the document for different versions of Microsoft Word.

    def test_optimize_for(self):
        doc = aw.Document()
        # This object contains an extensive list of flags unique to each document
        # that allow us to facilitate backward compatibility with older versions of Microsoft Word.
        options = doc.compatibility_options
        # Print the default settings for a blank document.
        print('\nDefault optimization settings:')
        ExCompatibilityOptions._print_compatibility_options(options)
        # We can access these settings in Microsoft Word via "File" -> "Options" -> "Advanced" -> "Compatibility options for...".
        doc.save(file_name=ARTIFACTS_DIR + 'CompatibilityOptions.OptimizeFor.DefaultSettings.docx')
        # We can use the OptimizeFor method to ensure optimal compatibility with a specific Microsoft Word version.
        doc.compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2010)
        print('\nOptimized for Word 2010:')
        ExCompatibilityOptions._print_compatibility_options(options)
        doc.compatibility_options.optimize_for(aw.settings.MsWordVersion.WORD2000)
        print('\nOptimized for Word 2000:')
        ExCompatibilityOptions._print_compatibility_options(options)

    @staticmethod
    def _print_compatibility_options(options):
        enabled_options = []
        disabled_options = []
        ExCompatibilityOptions._add_option_name(options.adjust_line_height_in_table, 'AdjustLineHeightInTable', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.align_tables_row_by_row, 'AlignTablesRowByRow', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.allow_space_of_same_style_in_table, 'AllowSpaceOfSameStyleInTable', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.apply_breaking_rules, 'ApplyBreakingRules', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.auto_space_like_word95, 'AutoSpaceLikeWord95', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.autofit_to_first_fixed_width_cell, 'AutofitToFirstFixedWidthCell', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.balance_single_byte_double_byte_width, 'BalanceSingleByteDoubleByteWidth', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.cached_col_balance, 'CachedColBalance', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.conv_mail_merge_esc, 'ConvMailMergeEsc', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.disable_open_type_font_formatting_features, 'DisableOpenTypeFontFormattingFeatures', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.display_hangul_fixed_width, 'DisplayHangulFixedWidth', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.do_not_autofit_constrained_tables, 'DoNotAutofitConstrainedTables', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.do_not_break_constrained_forced_table, 'DoNotBreakConstrainedForcedTable', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.do_not_break_wrapped_tables, 'DoNotBreakWrappedTables', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.do_not_expand_shift_return, 'DoNotExpandShiftReturn', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.do_not_leave_backslash_alone, 'DoNotLeaveBackslashAlone', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.do_not_snap_to_grid_in_cell, 'DoNotSnapToGridInCell', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.do_not_suppress_indentation, 'DoNotSnapToGridInCell', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.do_not_suppress_paragraph_borders, 'DoNotSuppressParagraphBorders', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.do_not_use_east_asian_break_rules, 'DoNotUseEastAsianBreakRules', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.do_not_use_html_paragraph_auto_spacing, 'DoNotUseHTMLParagraphAutoSpacing', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.do_not_use_indent_as_numbering_tab_stop, 'DoNotUseIndentAsNumberingTabStop', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.do_not_vert_align_cell_with_sp, 'DoNotVertAlignCellWithSp', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.do_not_vert_align_in_txbx, 'DoNotVertAlignInTxbx', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.do_not_wrap_text_with_punct, 'DoNotWrapTextWithPunct', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.footnote_layout_like_ww8, 'FootnoteLayoutLikeWW8', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.forget_last_tab_alignment, 'ForgetLastTabAlignment', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.grow_autofit, 'GrowAutofit', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.layout_raw_table_width, 'LayoutRawTableWidth', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.layout_table_rows_apart, 'LayoutTableRowsApart', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.line_wrap_like_word6, 'LineWrapLikeWord6', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.mw_small_caps, 'MWSmallCaps', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.no_column_balance, 'NoColumnBalance', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.no_extra_line_spacing, 'NoExtraLineSpacing', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.no_leading, 'NoLeading', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.no_space_raise_lower, 'NoSpaceRaiseLower', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.no_tab_hang_ind, 'NoTabHangInd', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.override_table_style_font_size_and_justification, 'OverrideTableStyleFontSizeAndJustification', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.print_body_text_before_header, 'PrintBodyTextBeforeHeader', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.print_col_black, 'PrintColBlack', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.select_fld_with_first_or_last_char, 'SelectFldWithFirstOrLastChar', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.shape_layout_like_ww8, 'ShapeLayoutLikeWW8', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.show_breaks_in_frames, 'ShowBreaksInFrames', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.space_for_ul, 'SpaceForUL', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.spacing_in_whole_points, 'SpacingInWholePoints', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.split_pg_break_and_para_mark, 'SplitPgBreakAndParaMark', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.sub_font_by_size, 'SubFontBySize', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.suppress_bottom_spacing, 'SuppressBottomSpacing', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.suppress_sp_bf_after_pg_brk, 'SuppressSpBfAfterPgBrk', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.suppress_spacing_at_top_of_page, 'SuppressSpacingAtTopOfPage', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.suppress_top_spacing, 'SuppressTopSpacing', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.suppress_top_spacing_wp, 'SuppressTopSpacingWP', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.swap_borders_facing_pgs, 'SwapBordersFacingPgs', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.swap_inside_and_outside_for_mirror_indents_and_relative_positioning, 'SwapInsideAndOutsideForMirrorIndentsAndRelativePositioning', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.transparent_metafiles, 'TransparentMetafiles', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.truncate_font_heights_like_wp6, 'TruncateFontHeightsLikeWP6', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.ui_compat_97_to_2003, 'UICompat97To2003', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.ul_trail_space, 'UlTrailSpace', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.underline_tab_in_num_list, 'UnderlineTabInNumList', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.use_alt_kinsoku_line_break_rules, 'UseAltKinsokuLineBreakRules', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.use_ansi_kerning_pairs, 'UseAnsiKerningPairs', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.use_fe_layout, 'UseFELayout', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.use_normal_style_for_list, 'UseNormalStyleForList', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.use_printer_metrics, 'UsePrinterMetrics', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.use_single_borderfor_contiguous_cells, 'UseSingleBorderforContiguousCells', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.use_word2002_table_style_rules, 'UseWord2002TableStyleRules', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.use_word2010_table_style_rules, 'UseWord2010TableStyleRules', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.use_word97_line_break_rules, 'UseWord97LineBreakRules', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.wp_justification, 'WPJustification', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.wp_space_width, 'WPSpaceWidth', enabled_options, disabled_options)
        ExCompatibilityOptions._add_option_name(options.wrap_trail_spaces, 'WrapTrailSpaces', enabled_options, disabled_options)
        print('\tEnabled options:')
        for option_name in enabled_options:
            print(f'\t\t{option_name}')
        print('\tDisabled options:')
        for option_name in disabled_options:
            print(f'\t\t{option_name}')

    @staticmethod
    def _add_option_name(option, option_name, enabled_options, disabled_options):
        if option:
            enabled_options.append(option_name)
        else:
            disabled_options.append(option_name)
    #ExEnd

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