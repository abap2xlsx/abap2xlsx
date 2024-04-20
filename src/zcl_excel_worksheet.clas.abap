CLASS zcl_excel_worksheet DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.
*"* public components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!

    INTERFACES zif_excel_sheet_printsettings .
    INTERFACES zif_excel_sheet_properties .
    INTERFACES zif_excel_sheet_protection .
    INTERFACES zif_excel_sheet_vba_project .

    TYPES:
      BEGIN OF  mty_s_outline_row,
        row_from  TYPE i,
        row_to    TYPE i,
        collapsed TYPE abap_bool,
      END OF mty_s_outline_row .
    TYPES:
      mty_ts_outlines_row TYPE SORTED TABLE OF mty_s_outline_row WITH UNIQUE KEY row_from row_to .
    TYPES:
      BEGIN OF mty_s_ignored_errors,
        "! Cell reference (e.g. "A1") or list like "A1 A2" or range "A1:G1"
        cell_coords           TYPE zexcel_cell_coords,
        "! Ignore errors when cells contain formulas that result in an error.
        eval_error            TYPE abap_bool,
        "! Ignore errors when formulas contain text formatted cells with years represented as 2 digits.
        two_digit_text_year   TYPE abap_bool,
        "! Ignore errors when numbers are formatted as text or are preceded by an apostrophe.
        number_stored_as_text TYPE abap_bool,
        "! Ignore errors when a formula in a region of your WorkSheet differs from other formulas in the same region.
        formula               TYPE abap_bool,
        "! Ignore errors when formulas omit certain cells in a region.
        formula_range         TYPE abap_bool,
        "! Ignore errors when unlocked cells contain formulas.
        unlocked_formula      TYPE abap_bool,
        "! Ignore errors when formulas refer to empty cells.
        empty_cell_reference  TYPE abap_bool,
        "! Ignore errors when a cell's value in a Table does not comply with the Data Validation rules specified.
        list_data_validation  TYPE abap_bool,
        "! Ignore errors when cells contain a value different from a calculated column formula.
        "! In other words, for a calculated column, a cell in that column is considered to have an error
        "! if its formula is different from the calculated column formula, or doesn't contain a formula at all.
        calculated_column     TYPE abap_bool,
      END OF mty_s_ignored_errors .
    TYPES:
      mty_th_ignored_errors TYPE HASHED TABLE OF mty_s_ignored_errors WITH UNIQUE KEY cell_coords .
    TYPES:
      BEGIN OF mty_s_column_formula,
        id                     TYPE i,
        column                 TYPE zexcel_cell_column,
        formula                TYPE string,
        table_top_left_row     TYPE zexcel_cell_row,
        table_bottom_right_row TYPE zexcel_cell_row,
        table_left_column_int  TYPE zexcel_cell_column,
        table_right_column_int TYPE zexcel_cell_column,
      END OF mty_s_column_formula .
    TYPES:
      mty_th_column_formula
               TYPE HASHED TABLE OF mty_s_column_formula
               WITH UNIQUE KEY id .
    TYPES:
      ty_doc_url TYPE c LENGTH 255 .
    TYPES:
      BEGIN OF mty_merge,
        row_from TYPE i,
        row_to   TYPE i,
        col_from TYPE i,
        col_to   TYPE i,
      END OF mty_merge .
    TYPES:
      mty_ts_merge TYPE SORTED TABLE OF mty_merge WITH UNIQUE KEY table_line .
    TYPES:
      ty_area TYPE c LENGTH 1 .

    CONSTANTS c_break_column TYPE zexcel_break VALUE 2 ##NO_TEXT.
    CONSTANTS c_break_none TYPE zexcel_break VALUE 0 ##NO_TEXT.
    CONSTANTS c_break_row TYPE zexcel_break VALUE 1 ##NO_TEXT.
    CONSTANTS:
      BEGIN OF c_area,
        whole   TYPE ty_area VALUE 'W',                     "#EC NOTEXT
        topleft TYPE ty_area VALUE 'T',                     "#EC NOTEXT
      END OF c_area .
    DATA excel TYPE REF TO zcl_excel READ-ONLY .
    DATA print_gridlines TYPE zexcel_print_gridlines READ-ONLY VALUE abap_false ##NO_TEXT.
    DATA sheet_content TYPE zexcel_t_cell_data .
    DATA sheet_setup TYPE REF TO zcl_excel_sheet_setup .
    DATA show_gridlines TYPE zexcel_show_gridlines READ-ONLY VALUE abap_true ##NO_TEXT.
    DATA show_rowcolheaders TYPE zexcel_show_gridlines READ-ONLY VALUE abap_true ##NO_TEXT.
    DATA tabcolor TYPE zexcel_s_tabcolor READ-ONLY .
    DATA column_formulas TYPE mty_th_column_formula READ-ONLY .
    CLASS-DATA:
      BEGIN OF c_messages READ-ONLY,
        formula_id_only_is_possible TYPE string,
        column_formula_id_not_found TYPE string,
        formula_not_in_this_table   TYPE string,
        formula_in_other_column     TYPE string,
      END OF c_messages .
    DATA mt_merged_cells TYPE mty_ts_merge READ-ONLY .
    DATA pane_top_left_cell TYPE string READ-ONLY.
    DATA sheetview_top_left_cell TYPE string READ-ONLY.

    METHODS add_comment
      IMPORTING
        !ip_comment TYPE REF TO zcl_excel_comment .
    METHODS add_drawing
      IMPORTING
        !ip_drawing TYPE REF TO zcl_excel_drawing .
    METHODS add_new_column
      IMPORTING
        !ip_column       TYPE simple
      RETURNING
        VALUE(eo_column) TYPE REF TO zcl_excel_column
      RAISING
        zcx_excel .
    METHODS add_new_style_cond
      IMPORTING
        !ip_dimension_range  TYPE string DEFAULT 'A1'
      RETURNING
        VALUE(eo_style_cond) TYPE REF TO zcl_excel_style_cond .
    METHODS add_new_data_validation
      RETURNING
        VALUE(eo_data_validation) TYPE REF TO zcl_excel_data_validation .
    METHODS add_new_range
      RETURNING
        VALUE(eo_range) TYPE REF TO zcl_excel_range .
    METHODS add_new_row
      IMPORTING
        !ip_row       TYPE simple
      RETURNING
        VALUE(eo_row) TYPE REF TO zcl_excel_row .
    METHODS bind_alv
      IMPORTING
        !io_alv      TYPE REF TO object
        !it_table    TYPE STANDARD TABLE
        !i_top       TYPE i DEFAULT 1
        !i_left      TYPE i DEFAULT 1
        !table_style TYPE zexcel_table_style OPTIONAL
        !i_table     TYPE abap_bool DEFAULT abap_true
      RAISING
        zcx_excel .
    METHODS bind_alv_ole2
      IMPORTING
        !i_document_url      TYPE ty_doc_url DEFAULT space
        !i_xls               TYPE c DEFAULT space
        !i_save_path         TYPE string
        !io_alv              TYPE REF TO cl_gui_alv_grid
        !it_listheader       TYPE slis_t_listheader OPTIONAL
        !i_top               TYPE i DEFAULT 1
        !i_left              TYPE i DEFAULT 1
        !i_columns_header    TYPE c DEFAULT 'X'
        !i_columns_autofit   TYPE c DEFAULT 'X'
        !i_format_col_header TYPE soi_format_item OPTIONAL
        !i_format_subtotal   TYPE soi_format_item OPTIONAL
        !i_format_total      TYPE soi_format_item OPTIONAL
      EXCEPTIONS
        miss_guide
        ex_transfer_kkblo_error
        fatal_error
        inv_data_range
        dim_mismatch_vkey
        dim_mismatch_sema
        error_in_sema .
    METHODS bind_table
      IMPORTING
        !ip_table               TYPE STANDARD TABLE
        !it_field_catalog       TYPE zexcel_t_fieldcatalog OPTIONAL
        !is_table_settings      TYPE zexcel_s_table_settings OPTIONAL
        VALUE(iv_default_descr) TYPE c OPTIONAL
        !iv_no_line_if_empty    TYPE abap_bool DEFAULT abap_false
        !ip_conv_exit_length    TYPE abap_bool DEFAULT abap_false
        !ip_conv_curr_amt_ext   TYPE abap_bool DEFAULT abap_false
      EXPORTING
        !es_table_settings      TYPE zexcel_s_table_settings
      RAISING
        zcx_excel .
    METHODS calculate_column_widths
      RAISING
        zcx_excel .
    METHODS change_area_style
      IMPORTING
        !ip_range         TYPE csequence OPTIONAL
        !ip_column_start  TYPE simple OPTIONAL
        !ip_column_end    TYPE simple OPTIONAL
        !ip_row           TYPE zexcel_cell_row OPTIONAL
        !ip_row_to        TYPE zexcel_cell_row OPTIONAL
        !ip_style_changer TYPE REF TO zif_excel_style_changer
      RAISING
        zcx_excel .
    METHODS change_cell_style
      IMPORTING
        !ip_columnrow                   TYPE csequence OPTIONAL
        !ip_column                      TYPE simple OPTIONAL
        !ip_row                         TYPE zexcel_cell_row OPTIONAL
        !ip_complete                    TYPE zexcel_s_cstyle_complete OPTIONAL
        !ip_xcomplete                   TYPE zexcel_s_cstylex_complete OPTIONAL
        !ip_font                        TYPE zexcel_s_cstyle_font OPTIONAL
        !ip_xfont                       TYPE zexcel_s_cstylex_font OPTIONAL
        !ip_fill                        TYPE zexcel_s_cstyle_fill OPTIONAL
        !ip_xfill                       TYPE zexcel_s_cstylex_fill OPTIONAL
        !ip_borders                     TYPE zexcel_s_cstyle_borders OPTIONAL
        !ip_xborders                    TYPE zexcel_s_cstylex_borders OPTIONAL
        !ip_alignment                   TYPE zexcel_s_cstyle_alignment OPTIONAL
        !ip_xalignment                  TYPE zexcel_s_cstylex_alignment OPTIONAL
        !ip_number_format_format_code   TYPE zexcel_number_format OPTIONAL
        !ip_protection                  TYPE zexcel_s_cstyle_protection OPTIONAL
        !ip_xprotection                 TYPE zexcel_s_cstylex_protection OPTIONAL
        !ip_font_bold                   TYPE flag OPTIONAL
        !ip_font_color                  TYPE zexcel_s_style_color OPTIONAL
        !ip_font_color_rgb              TYPE zexcel_style_color_argb OPTIONAL
        !ip_font_color_indexed          TYPE zexcel_style_color_indexed OPTIONAL
        !ip_font_color_theme            TYPE zexcel_style_color_theme OPTIONAL
        !ip_font_color_tint             TYPE zexcel_style_color_tint OPTIONAL
        !ip_font_family                 TYPE zexcel_style_font_family OPTIONAL
        !ip_font_italic                 TYPE flag OPTIONAL
        !ip_font_name                   TYPE zexcel_style_font_name OPTIONAL
        !ip_font_scheme                 TYPE zexcel_style_font_scheme OPTIONAL
        !ip_font_size                   TYPE zexcel_style_font_size OPTIONAL
        !ip_font_strikethrough          TYPE flag OPTIONAL
        !ip_font_underline              TYPE flag OPTIONAL
        !ip_font_underline_mode         TYPE zexcel_style_font_underline OPTIONAL
        !ip_fill_filltype               TYPE zexcel_fill_type OPTIONAL
        !ip_fill_rotation               TYPE zexcel_rotation OPTIONAL
        !ip_fill_fgcolor                TYPE zexcel_s_style_color OPTIONAL
        !ip_fill_fgcolor_rgb            TYPE zexcel_style_color_argb OPTIONAL
        !ip_fill_fgcolor_indexed        TYPE zexcel_style_color_indexed OPTIONAL
        !ip_fill_fgcolor_theme          TYPE zexcel_style_color_theme OPTIONAL
        !ip_fill_fgcolor_tint           TYPE zexcel_style_color_tint OPTIONAL
        !ip_fill_bgcolor                TYPE zexcel_s_style_color OPTIONAL
        !ip_fill_bgcolor_rgb            TYPE zexcel_style_color_argb OPTIONAL
        !ip_fill_bgcolor_indexed        TYPE zexcel_style_color_indexed OPTIONAL
        !ip_fill_bgcolor_theme          TYPE zexcel_style_color_theme OPTIONAL
        !ip_fill_bgcolor_tint           TYPE zexcel_style_color_tint OPTIONAL
        !ip_borders_allborders          TYPE zexcel_s_cstyle_border OPTIONAL
        !ip_fill_gradtype_type          TYPE zexcel_s_gradient_type-type OPTIONAL
        !ip_fill_gradtype_degree        TYPE zexcel_s_gradient_type-degree OPTIONAL
        !ip_xborders_allborders         TYPE zexcel_s_cstylex_border OPTIONAL
        !ip_borders_diagonal            TYPE zexcel_s_cstyle_border OPTIONAL
        !ip_fill_gradtype_bottom        TYPE zexcel_s_gradient_type-bottom OPTIONAL
        !ip_fill_gradtype_top           TYPE zexcel_s_gradient_type-top OPTIONAL
        !ip_xborders_diagonal           TYPE zexcel_s_cstylex_border OPTIONAL
        !ip_borders_diagonal_mode       TYPE zexcel_diagonal OPTIONAL
        !ip_fill_gradtype_right         TYPE zexcel_s_gradient_type-right OPTIONAL
        !ip_borders_down                TYPE zexcel_s_cstyle_border OPTIONAL
        !ip_fill_gradtype_left          TYPE zexcel_s_gradient_type-left OPTIONAL
        !ip_fill_gradtype_position1     TYPE zexcel_s_gradient_type-position1 OPTIONAL
        !ip_xborders_down               TYPE zexcel_s_cstylex_border OPTIONAL
        !ip_borders_left                TYPE zexcel_s_cstyle_border OPTIONAL
        !ip_fill_gradtype_position2     TYPE zexcel_s_gradient_type-position2 OPTIONAL
        !ip_fill_gradtype_position3     TYPE zexcel_s_gradient_type-position3 OPTIONAL
        !ip_xborders_left               TYPE zexcel_s_cstylex_border OPTIONAL
        !ip_borders_right               TYPE zexcel_s_cstyle_border OPTIONAL
        !ip_xborders_right              TYPE zexcel_s_cstylex_border OPTIONAL
        !ip_borders_top                 TYPE zexcel_s_cstyle_border OPTIONAL
        !ip_xborders_top                TYPE zexcel_s_cstylex_border OPTIONAL
        !ip_alignment_horizontal        TYPE zexcel_alignment OPTIONAL
        !ip_alignment_vertical          TYPE zexcel_alignment OPTIONAL
        !ip_alignment_textrotation      TYPE zexcel_text_rotation OPTIONAL
        !ip_alignment_wraptext          TYPE flag OPTIONAL
        !ip_alignment_shrinktofit       TYPE flag OPTIONAL
        !ip_alignment_indent            TYPE zexcel_indent OPTIONAL
        !ip_protection_hidden           TYPE zexcel_cell_protection OPTIONAL
        !ip_protection_locked           TYPE zexcel_cell_protection OPTIONAL
        !ip_borders_allborders_style    TYPE zexcel_border OPTIONAL
        !ip_borders_allborders_color    TYPE zexcel_s_style_color OPTIONAL
        !ip_borders_allbo_color_rgb     TYPE zexcel_style_color_argb OPTIONAL
        !ip_borders_allbo_color_indexed TYPE zexcel_style_color_indexed OPTIONAL
        !ip_borders_allbo_color_theme   TYPE zexcel_style_color_theme OPTIONAL
        !ip_borders_allbo_color_tint    TYPE zexcel_style_color_tint OPTIONAL
        !ip_borders_diagonal_style      TYPE zexcel_border OPTIONAL
        !ip_borders_diagonal_color      TYPE zexcel_s_style_color OPTIONAL
        !ip_borders_diagonal_color_rgb  TYPE zexcel_style_color_argb OPTIONAL
        !ip_borders_diagonal_color_inde TYPE zexcel_style_color_indexed OPTIONAL
        !ip_borders_diagonal_color_them TYPE zexcel_style_color_theme OPTIONAL
        !ip_borders_diagonal_color_tint TYPE zexcel_style_color_tint OPTIONAL
        !ip_borders_down_style          TYPE zexcel_border OPTIONAL
        !ip_borders_down_color          TYPE zexcel_s_style_color OPTIONAL
        !ip_borders_down_color_rgb      TYPE zexcel_style_color_argb OPTIONAL
        !ip_borders_down_color_indexed  TYPE zexcel_style_color_indexed OPTIONAL
        !ip_borders_down_color_theme    TYPE zexcel_style_color_theme OPTIONAL
        !ip_borders_down_color_tint     TYPE zexcel_style_color_tint OPTIONAL
        !ip_borders_left_style          TYPE zexcel_border OPTIONAL
        !ip_borders_left_color          TYPE zexcel_s_style_color OPTIONAL
        !ip_borders_left_color_rgb      TYPE zexcel_style_color_argb OPTIONAL
        !ip_borders_left_color_indexed  TYPE zexcel_style_color_indexed OPTIONAL
        !ip_borders_left_color_theme    TYPE zexcel_style_color_theme OPTIONAL
        !ip_borders_left_color_tint     TYPE zexcel_style_color_tint OPTIONAL
        !ip_borders_right_style         TYPE zexcel_border OPTIONAL
        !ip_borders_right_color         TYPE zexcel_s_style_color OPTIONAL
        !ip_borders_right_color_rgb     TYPE zexcel_style_color_argb OPTIONAL
        !ip_borders_right_color_indexed TYPE zexcel_style_color_indexed OPTIONAL
        !ip_borders_right_color_theme   TYPE zexcel_style_color_theme OPTIONAL
        !ip_borders_right_color_tint    TYPE zexcel_style_color_tint OPTIONAL
        !ip_borders_top_style           TYPE zexcel_border OPTIONAL
        !ip_borders_top_color           TYPE zexcel_s_style_color OPTIONAL
        !ip_borders_top_color_rgb       TYPE zexcel_style_color_argb OPTIONAL
        !ip_borders_top_color_indexed   TYPE zexcel_style_color_indexed OPTIONAL
        !ip_borders_top_color_theme     TYPE zexcel_style_color_theme OPTIONAL
        !ip_borders_top_color_tint      TYPE zexcel_style_color_tint OPTIONAL
      RETURNING
        VALUE(ep_guid)                  TYPE zexcel_cell_style
      RAISING
        zcx_excel .
    CLASS-METHODS class_constructor .
    METHODS constructor
      IMPORTING
        !ip_excel TYPE REF TO zcl_excel
        !ip_title TYPE zexcel_sheet_title OPTIONAL
      RAISING
        zcx_excel .
    METHODS delete_merge
      IMPORTING
        !ip_cell_column TYPE simple OPTIONAL
        !ip_cell_row    TYPE zexcel_cell_row OPTIONAL
      RAISING
        zcx_excel .
    METHODS delete_row_outline
      IMPORTING
        !iv_row_from TYPE i
        !iv_row_to   TYPE i
      RAISING
        zcx_excel .
    METHODS freeze_panes
      IMPORTING
        !ip_num_columns TYPE i OPTIONAL
        !ip_num_rows    TYPE i OPTIONAL
      RAISING
        zcx_excel .
    METHODS get_active_cell
      RETURNING
        VALUE(ep_active_cell) TYPE string
      RAISING
        zcx_excel .
    METHODS get_cell
      IMPORTING
        !ip_columnrow TYPE csequence OPTIONAL
        !ip_column    TYPE simple OPTIONAL
        !ip_row       TYPE zexcel_cell_row OPTIONAL
      EXPORTING
        !ep_value     TYPE zexcel_cell_value
        !ep_rc        TYPE sysubrc
        !ep_style     TYPE REF TO zcl_excel_style
        !ep_guid      TYPE zexcel_cell_style
        !ep_formula   TYPE zexcel_cell_formula
        !et_rtf       TYPE zexcel_t_rtf
      RAISING
        zcx_excel .
    METHODS get_column
      IMPORTING
        !ip_column       TYPE simple
      RETURNING
        VALUE(eo_column) TYPE REF TO zcl_excel_column
      RAISING
        zcx_excel .
    METHODS get_columns
      RETURNING
        VALUE(eo_columns) TYPE REF TO zcl_excel_columns
      RAISING
        zcx_excel .
    METHODS get_columns_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator
      RAISING
        zcx_excel .
    METHODS get_style_cond_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_data_validations_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_data_validations_size
      RETURNING
        VALUE(ep_size) TYPE i .
    METHODS get_default_column
      RETURNING
        VALUE(eo_column) TYPE REF TO zcl_excel_column
      RAISING
        zcx_excel .
    METHODS get_default_excel_date_format
      RETURNING
        VALUE(ep_default_excel_date_format) TYPE zexcel_number_format .
    METHODS get_default_excel_time_format
      RETURNING
        VALUE(ep_default_excel_time_format) TYPE zexcel_number_format .
    METHODS get_default_row
      RETURNING
        VALUE(eo_row) TYPE REF TO zcl_excel_row .
    METHODS get_dimension_range
      RETURNING
        VALUE(ep_dimension_range) TYPE string
      RAISING
        zcx_excel .
    METHODS get_comments
      RETURNING
        VALUE(r_comments) TYPE REF TO zcl_excel_comments .
    METHODS get_drawings
      IMPORTING
        !ip_type          TYPE zexcel_drawing_type OPTIONAL
      RETURNING
        VALUE(r_drawings) TYPE REF TO zcl_excel_drawings .
    METHODS get_comments_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_drawings_iterator
      IMPORTING
        !ip_type           TYPE zexcel_drawing_type
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_freeze_cell
      EXPORTING
        !ep_row    TYPE zexcel_cell_row
        !ep_column TYPE zexcel_cell_column .
    METHODS get_guid
      RETURNING
        VALUE(ep_guid) TYPE sysuuid_x16 .
    METHODS get_highest_column
      RETURNING
        VALUE(r_highest_column) TYPE zexcel_cell_column
      RAISING
        zcx_excel .
    METHODS get_highest_row
      RETURNING
        VALUE(r_highest_row) TYPE int4
      RAISING
        zcx_excel .
    METHODS get_hyperlinks_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_hyperlinks_size
      RETURNING
        VALUE(ep_size) TYPE i .
    METHODS get_ignored_errors
      RETURNING
        VALUE(rt_ignored_errors) TYPE mty_th_ignored_errors .
    METHODS get_merge
      RETURNING
        VALUE(merge_range) TYPE string_table
      RAISING
        zcx_excel .
    METHODS get_pagebreaks
      RETURNING
        VALUE(ro_pagebreaks) TYPE REF TO zcl_excel_worksheet_pagebreaks
      RAISING
        zcx_excel .
    METHODS get_ranges_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_row
      IMPORTING
        !ip_row       TYPE int4
      RETURNING
        VALUE(eo_row) TYPE REF TO zcl_excel_row .
    METHODS get_rows
      RETURNING
        VALUE(eo_rows) TYPE REF TO zcl_excel_rows .
    METHODS get_rows_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_row_outlines
      RETURNING
        VALUE(rt_row_outlines) TYPE mty_ts_outlines_row .
    METHODS get_style_cond
      IMPORTING
        !ip_guid             TYPE zexcel_cell_style
      RETURNING
        VALUE(eo_style_cond) TYPE REF TO zcl_excel_style_cond .
    METHODS get_tabcolor
      RETURNING
        VALUE(ev_tabcolor) TYPE zexcel_s_tabcolor .
    METHODS get_tables_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO zcl_excel_collection_iterator .
    METHODS get_tables_size
      RETURNING
        VALUE(ep_size) TYPE i .
    METHODS get_title
      IMPORTING
        !ip_escaped     TYPE flag DEFAULT ''
      RETURNING
        VALUE(ep_title) TYPE zexcel_sheet_title .
    METHODS is_cell_merged
      IMPORTING
        !ip_column          TYPE simple
        !ip_row             TYPE zexcel_cell_row
      RETURNING
        VALUE(rp_is_merged) TYPE abap_bool
      RAISING
        zcx_excel .
    METHODS set_cell
      IMPORTING
        !ip_columnrow         TYPE csequence OPTIONAL
        !ip_column            TYPE simple OPTIONAL
        !ip_row               TYPE zexcel_cell_row OPTIONAL
        !ip_value             TYPE simple OPTIONAL
        !ip_formula           TYPE zexcel_cell_formula OPTIONAL
        !ip_style             TYPE any OPTIONAL
        !ip_hyperlink         TYPE REF TO zcl_excel_hyperlink OPTIONAL
        !ip_data_type         TYPE zexcel_cell_data_type OPTIONAL
        !ip_abap_type         TYPE abap_typekind OPTIONAL
        !ip_currency          TYPE waers_curc OPTIONAL
        !it_rtf               TYPE zexcel_t_rtf OPTIONAL
        !ip_column_formula_id TYPE mty_s_column_formula-id OPTIONAL
        !ip_conv_exit_length  TYPE abap_bool DEFAULT abap_false
      RAISING
        zcx_excel .
    METHODS set_cell_formula
      IMPORTING
        !ip_columnrow TYPE csequence OPTIONAL
        !ip_column    TYPE simple OPTIONAL
        !ip_row       TYPE zexcel_cell_row OPTIONAL
        !ip_formula   TYPE zexcel_cell_formula
      RAISING
        zcx_excel .
    METHODS set_cell_style
      IMPORTING
        !ip_columnrow TYPE csequence OPTIONAL
        !ip_column    TYPE simple OPTIONAL
        !ip_row       TYPE zexcel_cell_row OPTIONAL
        !ip_style     TYPE any
      RAISING
        zcx_excel .
    METHODS set_column_width
      IMPORTING
        !ip_column         TYPE simple
        !ip_width_fix      TYPE simple DEFAULT 0
        !ip_width_autosize TYPE flag DEFAULT 'X'
      RAISING
        zcx_excel .
    METHODS set_default_excel_date_format
      IMPORTING
        !ip_default_excel_date_format TYPE zexcel_number_format
      RAISING
        zcx_excel .
    METHODS set_ignored_errors
      IMPORTING
        !it_ignored_errors TYPE mty_th_ignored_errors .
    METHODS set_merge
      IMPORTING
        !ip_range        TYPE csequence OPTIONAL
        !ip_column_start TYPE simple OPTIONAL
        !ip_column_end   TYPE simple OPTIONAL
        !ip_row          TYPE zexcel_cell_row OPTIONAL
        !ip_row_to       TYPE zexcel_cell_row OPTIONAL
        !ip_style        TYPE any OPTIONAL
        !ip_value        TYPE simple OPTIONAL          "added parameter
        !ip_formula      TYPE zexcel_cell_formula OPTIONAL        "added parameter
      RAISING
        zcx_excel .
    METHODS set_pane_top_left_cell
      IMPORTING
        !iv_columnrow TYPE csequence
      RAISING
        zcx_excel.
    METHODS set_print_gridlines
      IMPORTING
        !i_print_gridlines TYPE zexcel_print_gridlines .
    METHODS set_row_height
      IMPORTING
        !ip_row        TYPE simple
        !ip_height_fix TYPE simple
      RAISING
        zcx_excel .
    METHODS set_row_outline
      IMPORTING
        !iv_row_from  TYPE i
        !iv_row_to    TYPE i
        !iv_collapsed TYPE abap_bool
      RAISING
        zcx_excel .
    METHODS set_sheetview_top_left_cell
      IMPORTING
        !iv_columnrow TYPE csequence
      RAISING
        zcx_excel.
    METHODS set_show_gridlines
      IMPORTING
        !i_show_gridlines TYPE zexcel_show_gridlines .
    METHODS set_show_rowcolheaders
      IMPORTING
        !i_show_rowcolheaders TYPE zexcel_show_rowcolheader .
    METHODS set_tabcolor
      IMPORTING
        !iv_tabcolor TYPE zexcel_s_tabcolor .
    METHODS set_table
      IMPORTING
        !ip_table           TYPE STANDARD TABLE
        !ip_hdr_style       TYPE any OPTIONAL
        !ip_body_style      TYPE any OPTIONAL
        !ip_table_title     TYPE string
        !ip_top_left_column TYPE zexcel_cell_column_alpha DEFAULT 'B'
        !ip_top_left_row    TYPE zexcel_cell_row DEFAULT 3
        !ip_transpose       TYPE abap_bool OPTIONAL
        !ip_no_header       TYPE abap_bool OPTIONAL
      RAISING
        zcx_excel .
    METHODS set_title
      IMPORTING
        !ip_title TYPE zexcel_sheet_title
      RAISING
        zcx_excel .
    METHODS get_table
      IMPORTING
        !iv_skipped_rows           TYPE int4 DEFAULT 0
        !iv_skipped_cols           TYPE int4 DEFAULT 0
        !iv_max_col                TYPE int4 OPTIONAL
        !iv_max_row                TYPE int4 OPTIONAL
        !iv_skip_bottom_empty_rows TYPE abap_bool DEFAULT abap_false
      EXPORTING
        !et_table                  TYPE STANDARD TABLE
      RAISING
        zcx_excel .
    METHODS set_merge_style
      IMPORTING
        !ip_range        TYPE csequence OPTIONAL
        !ip_column_start TYPE simple OPTIONAL
        !ip_column_end   TYPE simple OPTIONAL
        !ip_row          TYPE zexcel_cell_row OPTIONAL
        !ip_row_to       TYPE zexcel_cell_row OPTIONAL
        !ip_style        TYPE any OPTIONAL
      RAISING
        zcx_excel .
    METHODS set_area_formula
      IMPORTING
        !ip_range        TYPE csequence OPTIONAL
        !ip_column_start TYPE simple OPTIONAL
        !ip_column_end   TYPE simple OPTIONAL
        !ip_row          TYPE zexcel_cell_row OPTIONAL
        !ip_row_to       TYPE zexcel_cell_row OPTIONAL
        !ip_formula      TYPE zexcel_cell_formula
        !ip_merge        TYPE abap_bool OPTIONAL
        !ip_area         TYPE ty_area DEFAULT c_area-topleft
      RAISING
        zcx_excel .
    METHODS set_area_style
      IMPORTING
        !ip_range        TYPE csequence OPTIONAL
        !ip_column_start TYPE simple OPTIONAL
        !ip_column_end   TYPE simple OPTIONAL
        !ip_row          TYPE zexcel_cell_row OPTIONAL
        !ip_row_to       TYPE zexcel_cell_row OPTIONAL
        !ip_style        TYPE any
        !ip_merge        TYPE abap_bool OPTIONAL
      RAISING
        zcx_excel .
    METHODS set_area
      IMPORTING
        !ip_range        TYPE csequence OPTIONAL
        !ip_column_start TYPE simple OPTIONAL
        !ip_column_end   TYPE simple OPTIONAL
        !ip_row          TYPE zexcel_cell_row OPTIONAL
        !ip_row_to       TYPE zexcel_cell_row OPTIONAL
        !ip_value        TYPE simple OPTIONAL
        !ip_formula      TYPE zexcel_cell_formula OPTIONAL
        !ip_style        TYPE any OPTIONAL
        !ip_hyperlink    TYPE REF TO zcl_excel_hyperlink OPTIONAL
        !ip_data_type    TYPE zexcel_cell_data_type OPTIONAL
        !ip_abap_type    TYPE abap_typekind OPTIONAL
        !ip_merge        TYPE abap_bool OPTIONAL
        !ip_area         TYPE ty_area DEFAULT c_area-topleft
      RAISING
        zcx_excel .
    METHODS get_header_footer_drawings
      RETURNING
        VALUE(rt_drawings) TYPE zexcel_t_drawings .
    METHODS set_area_hyperlink
      IMPORTING
        !ip_range        TYPE csequence OPTIONAL
        !ip_column_start TYPE simple OPTIONAL
        !ip_column_end   TYPE simple OPTIONAL
        !ip_row          TYPE zexcel_cell_row OPTIONAL
        !ip_row_to       TYPE zexcel_cell_row OPTIONAL
        !ip_url          TYPE string
        !ip_is_internal  TYPE abap_bool
      RAISING
        zcx_excel .
    "! excel upload, counterpart to BIND_TABLE
    "! @parameter it_field_catalog | field catalog, used to derive correct types
    "! @parameter iv_begin_row | starting row, by default 2 to skip header
    "! @parameter et_data | generic internal table, there may be conversion losses
    "! @parameter er_data | ref to internal table of string columns, to get raw data without conversion losses.
    METHODS convert_to_table
      IMPORTING
        !it_field_catalog TYPE zexcel_t_fieldcatalog OPTIONAL
        !iv_begin_row     TYPE int4 DEFAULT 2
        !iv_end_row       TYPE int4 DEFAULT 0
      EXPORTING
        !et_data          TYPE STANDARD TABLE
        !er_data          TYPE REF TO data
      RAISING
        zcx_excel .
  PROTECTED SECTION.
    METHODS set_table_reference
      IMPORTING
        !ip_column    TYPE zexcel_cell_column
        !ip_row       TYPE zexcel_cell_row
        !ir_table     TYPE REF TO zcl_excel_table
        !ip_fieldname TYPE zexcel_fieldname
        !ip_header    TYPE abap_bool
      RAISING
        zcx_excel .
  PRIVATE SECTION.

*"* private components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!
    TYPES ty_table_settings TYPE STANDARD TABLE OF zexcel_s_table_settings WITH DEFAULT KEY.
    DATA active_cell TYPE zexcel_s_cell_data .
    DATA charts TYPE REF TO zcl_excel_drawings .
    DATA columns TYPE REF TO zcl_excel_columns .
    DATA row_default TYPE REF TO zcl_excel_row .
    DATA column_default TYPE REF TO zcl_excel_column .
    DATA styles_cond TYPE REF TO zcl_excel_styles_cond .
    DATA data_validations TYPE REF TO zcl_excel_data_validations .
    DATA default_excel_date_format TYPE zexcel_number_format .
    DATA default_excel_time_format TYPE zexcel_number_format .
    DATA comments TYPE REF TO zcl_excel_comments .
    DATA drawings TYPE REF TO zcl_excel_drawings .
    DATA freeze_pane_cell_column TYPE zexcel_cell_column .
    DATA freeze_pane_cell_row TYPE zexcel_cell_row .
    DATA guid TYPE sysuuid_x16 .
    DATA hyperlinks TYPE REF TO zcl_excel_collection .
    DATA lower_cell TYPE zexcel_s_cell_data .
    DATA mo_pagebreaks TYPE REF TO zcl_excel_worksheet_pagebreaks .
    DATA mt_row_outlines TYPE mty_ts_outlines_row .
    DATA print_title_col_from TYPE zexcel_cell_column_alpha .
    DATA print_title_col_to TYPE zexcel_cell_column_alpha .
    DATA print_title_row_from TYPE zexcel_cell_row .
    DATA print_title_row_to TYPE zexcel_cell_row .
    DATA ranges TYPE REF TO zcl_excel_ranges .
    DATA rows TYPE REF TO zcl_excel_rows .
    DATA tables TYPE REF TO zcl_excel_collection .
    DATA title TYPE zexcel_sheet_title VALUE 'Worksheet'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    DATA upper_cell TYPE zexcel_s_cell_data .
    DATA mt_ignored_errors TYPE mty_th_ignored_errors.
    DATA right_to_left TYPE abap_bool.

    METHODS calculate_cell_width
      IMPORTING
        !ip_column      TYPE simple
        !ip_row         TYPE zexcel_cell_row
      RETURNING
        VALUE(ep_width) TYPE f
      RAISING
        zcx_excel .
    CLASS-METHODS calculate_table_bottom_right
      IMPORTING
        ip_table         TYPE STANDARD TABLE
        it_field_catalog TYPE zexcel_t_fieldcatalog
      CHANGING
        cs_settings      TYPE zexcel_s_table_settings
      RAISING
        zcx_excel.
    CLASS-METHODS check_cell_column_formula
      IMPORTING
        it_column_formulas   TYPE mty_th_column_formula
        ip_column_formula_id TYPE mty_s_column_formula-id
        ip_formula           TYPE zexcel_cell_formula
        ip_value             TYPE simple
        ip_row               TYPE zexcel_cell_row
        ip_column            TYPE zexcel_cell_column
      RAISING
        zcx_excel.
    METHODS check_rtf
      IMPORTING
        !ip_value       TYPE simple
        VALUE(ip_style) TYPE zexcel_cell_style OPTIONAL
      CHANGING
        !ct_rtf         TYPE zexcel_t_rtf
      RAISING
        zcx_excel .
    CLASS-METHODS check_table_overlapping
      IMPORTING
        is_table_settings       TYPE zexcel_s_table_settings
        it_other_table_settings TYPE ty_table_settings
      RAISING
        zcx_excel.
    METHODS clear_initial_colorxfields
      IMPORTING
        is_color  TYPE zexcel_s_style_color
      CHANGING
        cs_xcolor TYPE zexcel_s_cstylex_color.
    METHODS create_data_conv_exit_length
      IMPORTING
        !ip_value       TYPE simple
      RETURNING
        VALUE(ep_value) TYPE REF TO data.
    METHODS generate_title
      RETURNING
        VALUE(ep_title) TYPE zexcel_sheet_title .
    METHODS get_value_type
      IMPORTING
        !ip_value      TYPE simple
      EXPORTING
        !ep_value      TYPE simple
        !ep_value_type TYPE abap_typekind .
    METHODS move_supplied_borders
      IMPORTING
        iv_border_supplied        TYPE abap_bool
        is_border                 TYPE zexcel_s_cstyle_border
        iv_xborder_supplied       TYPE abap_bool
        is_xborder                TYPE zexcel_s_cstylex_border
      CHANGING
        cs_complete_style_border  TYPE zexcel_s_cstyle_border
        cs_complete_stylex_border TYPE zexcel_s_cstylex_border.
    METHODS normalize_column_heading_texts
      IMPORTING
        iv_default_descr TYPE c
        it_field_catalog TYPE zexcel_t_fieldcatalog
      RETURNING
        VALUE(result)    TYPE zexcel_t_fieldcatalog.
    METHODS normalize_columnrow_parameter
      IMPORTING
        ip_columnrow TYPE csequence OPTIONAL
        ip_column    TYPE simple OPTIONAL
        ip_row       TYPE zexcel_cell_row OPTIONAL
      EXPORTING
        ep_column    TYPE zexcel_cell_column
        ep_row       TYPE zexcel_cell_row
      RAISING
        zcx_excel.
    METHODS normalize_range_parameter
      IMPORTING
        ip_range        TYPE csequence OPTIONAL
        ip_column_start TYPE simple OPTIONAL
        ip_column_end   TYPE simple OPTIONAL
        ip_row          TYPE zexcel_cell_row OPTIONAL
        ip_row_to       TYPE zexcel_cell_row OPTIONAL
      EXPORTING
        ep_column_start TYPE zexcel_cell_column
        ep_column_end   TYPE zexcel_cell_column
        ep_row          TYPE zexcel_cell_row
        ep_row_to       TYPE zexcel_cell_row
      RAISING
        zcx_excel.
    CLASS-METHODS normalize_style_parameter
      IMPORTING
        !ip_style_or_guid TYPE any
      RETURNING
        VALUE(rv_guid)    TYPE zexcel_cell_style
      RAISING
        zcx_excel .
    METHODS print_title_set_range .
    METHODS update_dimension_range
      RAISING
        zcx_excel .
ENDCLASS.



CLASS zcl_excel_worksheet IMPLEMENTATION.


  METHOD add_comment.
    comments->include( ip_comment ).
  ENDMETHOD.                    "add_comment


  METHOD add_drawing.
    CASE ip_drawing->get_type( ).
      WHEN zcl_excel_drawing=>type_image.
        drawings->include( ip_drawing ).
      WHEN zcl_excel_drawing=>type_chart.
        charts->include( ip_drawing ).
    ENDCASE.
  ENDMETHOD.                    "ADD_DRAWING


  METHOD add_new_column.
    DATA: lv_column_alpha TYPE zexcel_cell_column_alpha.

    lv_column_alpha = zcl_excel_common=>convert_column2alpha( ip_column ).

    CREATE OBJECT eo_column
      EXPORTING
        ip_index     = lv_column_alpha
        ip_excel     = me->excel
        ip_worksheet = me.
    columns->add( eo_column ).
  ENDMETHOD.                    "ADD_NEW_COLUMN


  METHOD add_new_data_validation.

    CREATE OBJECT eo_data_validation.
    data_validations->add( eo_data_validation ).
  ENDMETHOD.                    "ADD_NEW_DATA_VALIDATION


  METHOD add_new_range.
* Create default blank range
    CREATE OBJECT eo_range.
    ranges->add( eo_range ).
  ENDMETHOD.                    "ADD_NEW_RANGE


  METHOD add_new_row.
    CREATE OBJECT eo_row
      EXPORTING
        ip_index = ip_row.
    rows->add( eo_row ).
  ENDMETHOD.                    "ADD_NEW_ROW


  METHOD add_new_style_cond.
    CREATE OBJECT eo_style_cond EXPORTING ip_dimension_range = ip_dimension_range.
    styles_cond->add( eo_style_cond ).
  ENDMETHOD.                    "ADD_NEW_STYLE_COND


  METHOD bind_alv.
    DATA: lo_converter TYPE REF TO zcl_excel_converter.

    CREATE OBJECT lo_converter.

    TRY.
        lo_converter->convert(
          EXPORTING
            io_alv         = io_alv
            it_table       = it_table
            i_row_int      = i_top
            i_column_int   = i_left
            i_table        = i_table
            i_style_table  = table_style
            io_worksheet   = me
          CHANGING
            co_excel       = excel ).
      CATCH zcx_excel .
    ENDTRY.

  ENDMETHOD.                    "BIND_ALV


  METHOD bind_alv_ole2.

    CALL METHOD ('ZCL_EXCEL_OLE')=>('BIND_ALV_OLE2')
      EXPORTING
        i_document_url          = i_document_url
        i_xls                   = i_xls
        i_save_path             = i_save_path
        io_alv                  = io_alv
        it_listheader           = it_listheader
        i_top                   = i_top
        i_left                  = i_left
        i_columns_header        = i_columns_header
        i_columns_autofit       = i_columns_autofit
        i_format_col_header     = i_format_col_header
        i_format_subtotal       = i_format_subtotal
        i_format_total          = i_format_total
      EXCEPTIONS
        miss_guide              = 1
        ex_transfer_kkblo_error = 2
        fatal_error             = 3
        inv_data_range          = 4
        dim_mismatch_vkey       = 5
        dim_mismatch_sema       = 6
        error_in_sema           = 7
        OTHERS                  = 8.
    IF sy-subrc <> 0.
      CASE sy-subrc.
        WHEN 1. RAISE miss_guide.
        WHEN 2. RAISE ex_transfer_kkblo_error.
        WHEN 3. RAISE fatal_error.
        WHEN 4. RAISE inv_data_range.
        WHEN 5. RAISE dim_mismatch_vkey.
        WHEN 6. RAISE dim_mismatch_sema.
        WHEN 7. RAISE error_in_sema.
      ENDCASE.
    ENDIF.

  ENDMETHOD.                    "BIND_ALV_OLE2


  METHOD bind_table.
*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmöcker,      (wi p)              2012-12-01
*              - ...
*          aligning code
*          message made to support multilinguality
*--------------------------------------------------------------------*
* issue #237   - Check if overlapping areas exist
*              - Alessandro Iannacci                        2012-12-01
* changes:     - Added raise if overlaps are detected
*--------------------------------------------------------------------*

    CONSTANTS:
      lc_top_left_column TYPE zexcel_cell_column_alpha VALUE 'A',
      lc_top_left_row    TYPE zexcel_cell_row VALUE 1,
      lc_no_currency     TYPE waers_curc VALUE IS INITIAL.

    DATA:
      lv_row_int              TYPE zexcel_cell_row,
      lv_first_row            TYPE zexcel_cell_row,
      lv_last_row             TYPE zexcel_cell_row,
      lv_column_int           TYPE zexcel_cell_column,
      lv_column_alpha         TYPE zexcel_cell_column_alpha,
      lt_field_catalog        TYPE zexcel_t_fieldcatalog,
      lv_id                   TYPE i,
      lv_formula              TYPE string,
      ls_settings             TYPE zexcel_s_table_settings,
      lo_table                TYPE REF TO zcl_excel_table,
      lv_value_lowercase      TYPE string,
      lv_syindex              TYPE c LENGTH 3,
      lo_iterator             TYPE REF TO zcl_excel_collection_iterator,
      lo_style_cond           TYPE REF TO zcl_excel_style_cond,
      lo_curtable             TYPE REF TO zcl_excel_table,
      lt_other_table_settings TYPE ty_table_settings.
    DATA: ls_column_formula TYPE mty_s_column_formula,
          lv_mincol         TYPE i.

    FIELD-SYMBOLS:
      <ls_field_catalog>        TYPE zexcel_s_fieldcatalog,
      <ls_field_catalog_custom> TYPE zexcel_s_fieldcatalog,
      <fs_table_line>           TYPE any,
      <fs_fldval>               TYPE any,
      <fs_fldval_currency>      TYPE waers.

    ls_settings = is_table_settings.

    IF ls_settings-top_left_column IS INITIAL.
      ls_settings-top_left_column = lc_top_left_column.
    ENDIF.

    IF ls_settings-table_style IS INITIAL.
      ls_settings-table_style = zcl_excel_table=>builtinstyle_medium2.
    ENDIF.

    IF ls_settings-top_left_row IS INITIAL.
      ls_settings-top_left_row = lc_top_left_row.
    ENDIF.

    IF it_field_catalog IS NOT SUPPLIED.
      lt_field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = ip_table
                                                             ip_conv_exit_length = ip_conv_exit_length ).
    ELSE.
      lt_field_catalog = it_field_catalog.
    ENDIF.

    SORT lt_field_catalog BY position.

    calculate_table_bottom_right(
      EXPORTING
        ip_table         = ip_table
        it_field_catalog = lt_field_catalog
      CHANGING
        cs_settings      = ls_settings ).

* Check if overlapping areas exist

    lo_iterator = me->tables->get_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_curtable ?= lo_iterator->get_next( ).
      APPEND lo_curtable->settings TO lt_other_table_settings.
    ENDWHILE.

    check_table_overlapping(
        is_table_settings       = ls_settings
        it_other_table_settings = lt_other_table_settings ).

* Start filling the table

    CREATE OBJECT lo_table.
    lo_table->settings = ls_settings.
    lo_table->set_data( ir_data = ip_table ).
    lv_id = me->excel->get_next_table_id( ).
    lo_table->set_id( iv_id = lv_id ).

    me->tables->add( lo_table ).

    lv_column_int = zcl_excel_common=>convert_column2int( ls_settings-top_left_column ).
    lv_row_int = ls_settings-top_left_row.

    lt_field_catalog = normalize_column_heading_texts(
          iv_default_descr = iv_default_descr
          it_field_catalog = lt_field_catalog ).

* It is better to loop column by column (only visible column)
    LOOP AT lt_field_catalog ASSIGNING <ls_field_catalog> WHERE dynpfld EQ abap_true.

      lv_column_alpha = zcl_excel_common=>convert_column2alpha( lv_column_int ).

      " First of all write column header
      IF <ls_field_catalog>-style_header IS NOT INITIAL.
        me->set_cell( ip_column = lv_column_alpha
                      ip_row    = lv_row_int
                      ip_value  = <ls_field_catalog>-column_name
                      ip_style  = <ls_field_catalog>-style_header ).
      ELSE.
        me->set_cell( ip_column = lv_column_alpha
                      ip_row    = lv_row_int
                      ip_value  = <ls_field_catalog>-column_name ).
      ENDIF.

      me->set_table_reference( ip_column    = lv_column_int
                               ip_row       = lv_row_int
                               ir_table     = lo_table
                               ip_fieldname = <ls_field_catalog>-fieldname
                               ip_header    = abap_true ).

      IF <ls_field_catalog>-column_formula IS NOT INITIAL.
        ls_column_formula-id                     = lines( column_formulas ) + 1.
        ls_column_formula-column                 = lv_column_int.
        ls_column_formula-formula                = <ls_field_catalog>-column_formula.
        ls_column_formula-table_top_left_row     = lo_table->settings-top_left_row.
        ls_column_formula-table_bottom_right_row = lo_table->settings-bottom_right_row.
        ls_column_formula-table_left_column_int  = lv_mincol.
        ls_column_formula-table_right_column_int = zcl_excel_common=>convert_column2int( lo_table->settings-bottom_right_column ).
        INSERT ls_column_formula INTO TABLE column_formulas.
      ENDIF.

      ADD 1 TO lv_row_int.
      LOOP AT ip_table ASSIGNING <fs_table_line>.

        ASSIGN COMPONENT <ls_field_catalog>-fieldname OF STRUCTURE <fs_table_line> TO <fs_fldval>.

        " issue #290 Add formula support in table
        IF <ls_field_catalog>-formula EQ abap_true.
          IF <ls_field_catalog>-style IS NOT INITIAL.
            IF <ls_field_catalog>-abap_type IS NOT INITIAL.
              me->set_cell( ip_column   = lv_column_alpha
                          ip_row      = lv_row_int
                          ip_formula  = <fs_fldval>
                          ip_abap_type = <ls_field_catalog>-abap_type
                          ip_style    = <ls_field_catalog>-style ).
            ELSE.
              me->set_cell( ip_column   = lv_column_alpha
                            ip_row      = lv_row_int
                            ip_formula  = <fs_fldval>
                            ip_style    = <ls_field_catalog>-style ).
            ENDIF.
          ELSEIF <ls_field_catalog>-abap_type IS NOT INITIAL.
            me->set_cell( ip_column   = lv_column_alpha
                          ip_row      = lv_row_int
                          ip_formula  = <fs_fldval>
                          ip_abap_type = <ls_field_catalog>-abap_type ).
          ELSE.
            me->set_cell( ip_column   = lv_column_alpha
                          ip_row      = lv_row_int
                          ip_formula  = <fs_fldval> ).
          ENDIF.
        ELSEIF <ls_field_catalog>-column_formula IS NOT INITIAL.
          " Column formulas
          IF <ls_field_catalog>-style IS NOT INITIAL.
            IF <ls_field_catalog>-abap_type IS NOT INITIAL.
              me->set_cell( ip_column            = lv_column_alpha
                            ip_row               = lv_row_int
                            ip_column_formula_id = ls_column_formula-id
                            ip_abap_type         = <ls_field_catalog>-abap_type
                            ip_style             = <ls_field_catalog>-style ).
            ELSE.
              me->set_cell( ip_column            = lv_column_alpha
                            ip_row               = lv_row_int
                            ip_column_formula_id = ls_column_formula-id
                            ip_style             = <ls_field_catalog>-style ).
            ENDIF.
          ELSEIF <ls_field_catalog>-abap_type IS NOT INITIAL.
            me->set_cell( ip_column             = lv_column_alpha
                          ip_row                = lv_row_int
                          ip_column_formula_id  = ls_column_formula-id
                          ip_abap_type          = <ls_field_catalog>-abap_type ).
          ELSE.
            me->set_cell( ip_column            = lv_column_alpha
                          ip_row               = lv_row_int
                          ip_column_formula_id = ls_column_formula-id ).
          ENDIF.
        ELSE.
          IF <ls_field_catalog>-currency_column IS INITIAL OR ip_conv_curr_amt_ext = abap_false.
            ASSIGN lc_no_currency TO <fs_fldval_currency>.
          ELSE.
            ASSIGN COMPONENT <ls_field_catalog>-currency_column OF STRUCTURE <fs_table_line> TO <fs_fldval_currency>.
          ENDIF.

          IF <ls_field_catalog>-style IS NOT INITIAL.
            IF <ls_field_catalog>-abap_type IS NOT INITIAL.
              me->set_cell( ip_column           = lv_column_alpha
                            ip_row              = lv_row_int
                            ip_value            = <fs_fldval>
                            ip_abap_type        = <ls_field_catalog>-abap_type
                            ip_currency         = <fs_fldval_currency>
                            ip_style            = <ls_field_catalog>-style
                            ip_conv_exit_length = ip_conv_exit_length ).
            ELSE.
              me->set_cell( ip_column = lv_column_alpha
                            ip_row    = lv_row_int
                            ip_value  = <fs_fldval>
                            ip_currency = <fs_fldval_currency>
                            ip_style  = <ls_field_catalog>-style
                            ip_conv_exit_length = ip_conv_exit_length ).
            ENDIF.
          ELSE.
            IF <ls_field_catalog>-abap_type IS NOT INITIAL.
              me->set_cell( ip_column = lv_column_alpha
                          ip_row    = lv_row_int
                          ip_abap_type = <ls_field_catalog>-abap_type
                          ip_currency  = <fs_fldval_currency>
                          ip_value  = <fs_fldval>
                          ip_conv_exit_length = ip_conv_exit_length ).
            ELSE.
              me->set_cell( ip_column = lv_column_alpha
                            ip_row    = lv_row_int
                            ip_currency = <fs_fldval_currency>
                            ip_value  = <fs_fldval>
                            ip_conv_exit_length = ip_conv_exit_length ).
            ENDIF.
          ENDIF.
        ENDIF.
        ADD 1 TO lv_row_int.

      ENDLOOP.
      IF sy-subrc <> 0 AND iv_no_line_if_empty = abap_false. "create empty row if table has no data
        me->set_cell( ip_column = lv_column_alpha
                      ip_row    = lv_row_int
                      ip_value  = space ).
        ADD 1 TO lv_row_int.
      ENDIF.

*--------------------------------------------------------------------*
      " totals
*--------------------------------------------------------------------*
      IF <ls_field_catalog>-totals_function IS NOT INITIAL.
        lv_formula = lo_table->get_totals_formula( ip_column = <ls_field_catalog>-column_name ip_function = <ls_field_catalog>-totals_function ).
        IF <ls_field_catalog>-style_total IS NOT INITIAL.
          me->set_cell( ip_column   = lv_column_alpha
                        ip_row      = lv_row_int
                        ip_formula  = lv_formula
                        ip_style    = <ls_field_catalog>-style_total ).
        ELSE.
          me->set_cell( ip_column   = lv_column_alpha
                        ip_row      = lv_row_int
                        ip_formula  = lv_formula ).
        ENDIF.
      ENDIF.

      lv_row_int = ls_settings-top_left_row.
      ADD 1 TO lv_column_int.

*--------------------------------------------------------------------*
      " conditional formatting
*--------------------------------------------------------------------*
      IF <ls_field_catalog>-style_cond IS NOT INITIAL.
        lv_first_row    = ls_settings-top_left_row + 1. " +1 to exclude header
        lv_last_row     = ls_settings-bottom_right_row.
        lo_style_cond = me->get_style_cond( <ls_field_catalog>-style_cond ).
        lo_style_cond->set_range( ip_start_column  = lv_column_alpha
                                  ip_start_row     = lv_first_row
                                  ip_stop_column   = lv_column_alpha
                                  ip_stop_row      = lv_last_row ).
      ENDIF.

    ENDLOOP.

*--------------------------------------------------------------------*
    " Set field catalog
*--------------------------------------------------------------------*
    lo_table->fieldcat = lt_field_catalog[].

    es_table_settings = ls_settings.
    es_table_settings-bottom_right_column = lv_column_alpha.
    " >> Issue #291
    IF ip_table IS INITIAL.
      es_table_settings-bottom_right_row    = ls_settings-top_left_row + 2.           "Last rows
    ELSE.
      es_table_settings-bottom_right_row    = ls_settings-bottom_right_row + 1. "Last rows
    ENDIF.
    " << Issue #291

  ENDMETHOD.                    "BIND_TABLE


  METHOD calculate_cell_width.
*--------------------------------------------------------------------*
* issue #293   - Roberto Bianco
*              - Christian Assig                            2014-03-14
*
* changes: - Calculate widths using SAPscript font metrics
*            (transaction SE73)
*          - Calculate the width of dates
*          - Add additional width for auto filter buttons
*          - Add cell padding to simulate Excel behavior
*--------------------------------------------------------------------*

    DATA: ld_cell_value                TYPE zexcel_cell_value,
          ld_style_guid                TYPE zexcel_cell_style,
          ls_stylemapping              TYPE zexcel_s_stylemapping,
          lo_table_object              TYPE REF TO object,
          lo_table                     TYPE REF TO zcl_excel_table,
          ld_table_top_left_column     TYPE zexcel_cell_column,
          ld_table_bottom_right_column TYPE zexcel_cell_column,
          ld_flag_contains_auto_filter TYPE abap_bool VALUE abap_false,
          ld_flag_bold                 TYPE abap_bool VALUE abap_false,
          ld_flag_italic               TYPE abap_bool VALUE abap_false,
          ld_date                      TYPE d,
          ld_date_char                 TYPE c LENGTH 50,
          ld_font_height               TYPE tdfontsize VALUE zcl_excel_font=>lc_default_font_height,
          ld_font_name                 TYPE zexcel_style_font_name VALUE zcl_excel_font=>lc_default_font_name.

    " Determine cell content and cell style
    me->get_cell( EXPORTING ip_column = ip_column
                            ip_row    = ip_row
                  IMPORTING ep_value  = ld_cell_value
                            ep_guid   = ld_style_guid ).

    " ABAP2XLSX uses tables to define areas containing headers and
    " auto-filters. Find out if the current cell is in the header
    " of one of these tables.
    LOOP AT me->tables->collection INTO lo_table_object.
      " Downcast: OBJECT -> ZCL_EXCEL_TABLE
      lo_table ?= lo_table_object.

      " Convert column letters to corresponding integer values
      ld_table_top_left_column =
        zcl_excel_common=>convert_column2int(
          lo_table->settings-top_left_column ).

      ld_table_bottom_right_column =
        zcl_excel_common=>convert_column2int(
          lo_table->settings-bottom_right_column ).

      " Is the current cell part of the table header?
      IF ip_column BETWEEN ld_table_top_left_column AND
                           ld_table_bottom_right_column AND
         ip_row    EQ lo_table->settings-top_left_row.
        " Current cell is part of the table header
        " -> Assume that an auto filter is present and that the font is
        "    bold
        ld_flag_contains_auto_filter = abap_true.
        ld_flag_bold = abap_true.
      ENDIF.
    ENDLOOP.

    " If a style GUID is present, read style attributes
    IF ld_style_guid IS NOT INITIAL.
      TRY.
          " Read style attributes
          ls_stylemapping = me->excel->get_style_to_guid( ld_style_guid ).

          " If the current cell contains the default date format,
          " convert the cell value to a date and calculate its length
          IF ls_stylemapping-complete_style-number_format-format_code =
             zcl_excel_style_number_format=>c_format_date_std.

            " Convert excel date to ABAP date
            ld_date =
              zcl_excel_common=>excel_string_to_date( ld_cell_value ).

            " Format ABAP date using user's formatting settings
            WRITE ld_date TO ld_date_char.

            " Remember the formatted date to calculate the cell size
            ld_cell_value = ld_date_char.

          ENDIF.

          " Read the font size and convert it to the font height
          " used by SAPscript (multiplication by 10)
          IF ls_stylemapping-complete_stylex-font-size = abap_true.
            ld_font_height = ls_stylemapping-complete_style-font-size * 10.
          ENDIF.

          " If set, remember the font name
          IF ls_stylemapping-complete_stylex-font-name = abap_true.
            ld_font_name = ls_stylemapping-complete_style-font-name.
          ENDIF.

          " If set, remember whether font is bold and italic.
          IF ls_stylemapping-complete_stylex-font-bold = abap_true.
            ld_flag_bold = ls_stylemapping-complete_style-font-bold.
          ENDIF.

          IF ls_stylemapping-complete_stylex-font-italic = abap_true.
            ld_flag_italic = ls_stylemapping-complete_style-font-italic.
          ENDIF.

        CATCH zcx_excel.                                "#EC NO_HANDLER
          " Style GUID is present, but style was not found
          " Continue with default values

      ENDTRY.
    ENDIF.

    ep_width = zcl_excel_font=>calculate_text_width(
      iv_font_name   = ld_font_name
      iv_font_height = ld_font_height
      iv_flag_bold   = ld_flag_bold
      iv_flag_italic = ld_flag_italic
      iv_cell_value  = ld_cell_value ).

    " If the current cell contains an auto filter, make it a bit wider.
    " The size used by the auto filter button does not depend on the font
    " size.
    IF ld_flag_contains_auto_filter = abap_true.
      ADD 2 TO ep_width.
    ENDIF.

  ENDMETHOD.


  METHOD calculate_column_widths.
    TYPES:
      BEGIN OF t_auto_size,
        col_index TYPE int4,
        width     TYPE f,
      END   OF t_auto_size.
    TYPES: tt_auto_size TYPE TABLE OF t_auto_size.

    DATA: lo_column_iterator TYPE REF TO zcl_excel_collection_iterator,
          lo_column          TYPE REF TO zcl_excel_column.

    DATA: auto_size   TYPE flag.
    DATA: auto_sizes  TYPE tt_auto_size.
    DATA: count       TYPE int4.
    DATA: highest_row TYPE int4.
    DATA: width       TYPE f.

    FIELD-SYMBOLS: <auto_size>        LIKE LINE OF auto_sizes.

    lo_column_iterator = me->get_columns_iterator( ).
    WHILE lo_column_iterator->has_next( ) = abap_true.
      lo_column ?= lo_column_iterator->get_next( ).
      auto_size = lo_column->get_auto_size( ).
      IF auto_size = abap_true.
        APPEND INITIAL LINE TO auto_sizes ASSIGNING <auto_size>.
        <auto_size>-col_index = lo_column->get_column_index( ).
        <auto_size>-width     = -1.
      ENDIF.
    ENDWHILE.

    " There is only something to do if there are some auto-size columns
    IF NOT auto_sizes IS INITIAL.
      highest_row = me->get_highest_row( ).
      LOOP AT auto_sizes ASSIGNING <auto_size>.
        count = 1.
        WHILE count <= highest_row.
* Do not check merged cells
          IF is_cell_merged(
              ip_column    = <auto_size>-col_index
              ip_row       = count ) = abap_false.
            width = calculate_cell_width( ip_column = <auto_size>-col_index     " issue #155 - less restrictive typing for ip_column
                                          ip_row    = count ).
            IF width > <auto_size>-width.
              <auto_size>-width = width.
            ENDIF.
          ENDIF.
          count = count + 1.
        ENDWHILE.
        lo_column = me->get_column( <auto_size>-col_index ). " issue #155 - less restrictive typing for ip_column
        lo_column->set_width( <auto_size>-width ).
      ENDLOOP.
    ENDIF.

  ENDMETHOD.                    "CALCULATE_COLUMN_WIDTHS


  METHOD calculate_table_bottom_right.

    DATA: lv_errormessage TYPE string,
          lv_columns      TYPE i,
          lt_columns      TYPE zexcel_t_fieldcatalog,
          lv_maxrow       TYPE i,
          lo_iterator     TYPE REF TO zcl_excel_collection_iterator,
          lo_curtable     TYPE REF TO zcl_excel_table,
          lv_row_int      TYPE zexcel_cell_row,
          lv_column_int   TYPE zexcel_cell_column,
          lv_rows         TYPE i,
          lv_maxcol       TYPE i.

    "Get the number of columns for the current table
    lt_columns = it_field_catalog.
    DELETE lt_columns WHERE dynpfld NE abap_true.
    DESCRIBE TABLE lt_columns LINES lv_columns.

    "Calculate the top left row of the current table
    lv_column_int = zcl_excel_common=>convert_column2int( cs_settings-top_left_column ).
    lv_row_int    = cs_settings-top_left_row.

    "Get number of row for the current table
    DESCRIBE TABLE ip_table LINES lv_rows.

    "Calculate the bottom right row for the current table
    lv_maxcol                       = lv_column_int + lv_columns - 1.
    lv_maxrow                       = lv_row_int    + lv_rows.
    cs_settings-bottom_right_column = zcl_excel_common=>convert_column2alpha( lv_maxcol ).
    cs_settings-bottom_right_row    = lv_maxrow.

  ENDMETHOD.


  METHOD change_area_style.

    DATA: lv_row              TYPE zexcel_cell_row,
          lv_row_start        TYPE zexcel_cell_row,
          lv_row_to           TYPE zexcel_cell_row,
          lv_column_int       TYPE zexcel_cell_column,
          lv_column_start_int TYPE zexcel_cell_column,
          lv_column_end_int   TYPE zexcel_cell_column.

    normalize_range_parameter( EXPORTING ip_range        = ip_range
                                         ip_column_start = ip_column_start     ip_column_end = ip_column_end
                                         ip_row          = ip_row              ip_row_to     = ip_row_to
                               IMPORTING ep_column_start = lv_column_start_int ep_column_end = lv_column_end_int
                                         ep_row          = lv_row_start        ep_row_to     = lv_row_to ).

    lv_column_int = lv_column_start_int.
    WHILE lv_column_int <= lv_column_end_int.

      lv_row = lv_row_start.
      WHILE lv_row <= lv_row_to.

        ip_style_changer->apply( ip_worksheet = me
                                 ip_column    = lv_column_int
                                 ip_row       = lv_row ).

        ADD 1 TO lv_row.
      ENDWHILE.

      ADD 1 TO lv_column_int.
    ENDWHILE.

  ENDMETHOD.


  METHOD change_cell_style.

    DATA: changer TYPE REF TO zif_excel_style_changer,
          column  TYPE zexcel_cell_column,
          row     TYPE zexcel_cell_row.

    normalize_columnrow_parameter( EXPORTING ip_columnrow = ip_columnrow
                                             ip_column    = ip_column
                                             ip_row       = ip_row
                                   IMPORTING ep_column    = column
                                             ep_row       = row ).

    changer = zcl_excel_style_changer=>create( excel = excel ).


    IF ip_complete IS SUPPLIED.
      IF ip_xcomplete IS NOT SUPPLIED.
        zcx_excel=>raise_text( 'Complete styleinfo has to be supplied with corresponding X-field' ).
      ENDIF.
      changer->set_complete( ip_complete = ip_complete ip_xcomplete = ip_xcomplete ).
    ENDIF.



    IF ip_font IS SUPPLIED.
      IF ip_xfont IS SUPPLIED.
        changer->set_complete_font( ip_font = ip_font ip_xfont = ip_xfont ).
      ELSE.
        changer->set_complete_font( ip_font = ip_font ).
      ENDIF.
    ENDIF.

    IF ip_fill IS SUPPLIED.
      IF ip_xfill IS SUPPLIED.
        changer->set_complete_fill( ip_fill = ip_fill ip_xfill = ip_xfill ).
      ELSE.
        changer->set_complete_fill( ip_fill = ip_fill ).
      ENDIF.
    ENDIF.


    IF ip_borders IS SUPPLIED.
      IF ip_xborders IS SUPPLIED.
        changer->set_complete_borders( ip_borders = ip_borders ip_xborders = ip_xborders ).
      ELSE.
        changer->set_complete_borders( ip_borders = ip_borders ).
      ENDIF.
    ENDIF.

    IF ip_alignment IS SUPPLIED.
      IF ip_xalignment IS SUPPLIED.
        changer->set_complete_alignment( ip_alignment = ip_alignment ip_xalignment = ip_xalignment ).
      ELSE.
        changer->set_complete_alignment( ip_alignment = ip_alignment ).
      ENDIF.
    ENDIF.

    IF ip_protection IS SUPPLIED.
      IF ip_xprotection IS SUPPLIED.
        changer->set_complete_protection( ip_protection = ip_protection ip_xprotection = ip_xprotection ).
      ELSE.
        changer->set_complete_protection( ip_protection = ip_protection ).
      ENDIF.
    ENDIF.


    IF ip_borders_allborders IS SUPPLIED.
      IF ip_xborders_allborders IS SUPPLIED.
        changer->set_complete_borders_all( ip_borders_allborders = ip_borders_allborders ip_xborders_allborders = ip_xborders_allborders ).
      ELSE.
        changer->set_complete_borders_all( ip_borders_allborders = ip_borders_allborders ).
      ENDIF.
    ENDIF.

    IF ip_borders_diagonal IS SUPPLIED.
      IF ip_xborders_diagonal IS SUPPLIED.
        changer->set_complete_borders_diagonal( ip_borders_diagonal = ip_borders_diagonal ip_xborders_diagonal = ip_xborders_diagonal ).
      ELSE.
        changer->set_complete_borders_diagonal( ip_borders_diagonal = ip_borders_diagonal ).
      ENDIF.
    ENDIF.

    IF ip_borders_down IS SUPPLIED.
      IF ip_xborders_down IS SUPPLIED.
        changer->set_complete_borders_down( ip_borders_down = ip_borders_down ip_xborders_down = ip_xborders_down ).
      ELSE.
        changer->set_complete_borders_down( ip_borders_down = ip_borders_down ).
      ENDIF.
    ENDIF.

    IF ip_borders_left IS SUPPLIED.
      IF ip_xborders_left IS SUPPLIED.
        changer->set_complete_borders_left( ip_borders_left = ip_borders_left ip_xborders_left = ip_xborders_left ).
      ELSE.
        changer->set_complete_borders_left( ip_borders_left = ip_borders_left ).
      ENDIF.
    ENDIF.

    IF ip_borders_right IS SUPPLIED.
      IF ip_xborders_right IS SUPPLIED.
        changer->set_complete_borders_right( ip_borders_right = ip_borders_right ip_xborders_right = ip_xborders_right ).
      ELSE.
        changer->set_complete_borders_right( ip_borders_right = ip_borders_right ).
      ENDIF.
    ENDIF.

    IF ip_borders_top IS SUPPLIED.
      IF ip_xborders_top IS SUPPLIED.
        changer->set_complete_borders_top( ip_borders_top = ip_borders_top ip_xborders_top = ip_xborders_top ).
      ELSE.
        changer->set_complete_borders_top( ip_borders_top = ip_borders_top ).
      ENDIF.
    ENDIF.

    IF ip_number_format_format_code IS SUPPLIED.
      changer->set_number_format( ip_number_format_format_code ).
    ENDIF.
    IF ip_font_bold IS SUPPLIED.
      changer->set_font_bold( ip_font_bold ).
    ENDIF.
    IF ip_font_color IS SUPPLIED.
      changer->set_font_color( ip_font_color ).
    ENDIF.
    IF ip_font_color_rgb IS SUPPLIED.
      changer->set_font_color_rgb( ip_font_color_rgb ).
    ENDIF.
    IF ip_font_color_indexed IS SUPPLIED.
      changer->set_font_color_indexed( ip_font_color_indexed ).
    ENDIF.
    IF ip_font_color_theme IS SUPPLIED.
      changer->set_font_color_theme( ip_font_color_theme ).
    ENDIF.
    IF ip_font_color_tint IS SUPPLIED.
      changer->set_font_color_tint( ip_font_color_tint ).
    ENDIF.

    IF ip_font_family IS SUPPLIED.
      changer->set_font_family( ip_font_family ).
    ENDIF.
    IF ip_font_italic IS SUPPLIED.
      changer->set_font_italic( ip_font_italic ).
    ENDIF.
    IF ip_font_name IS SUPPLIED.
      changer->set_font_name( ip_font_name ).
    ENDIF.
    IF ip_font_scheme IS SUPPLIED.
      changer->set_font_scheme( ip_font_scheme ).
    ENDIF.
    IF ip_font_size IS SUPPLIED.
      changer->set_font_size( ip_font_size ).
    ENDIF.
    IF ip_font_strikethrough IS SUPPLIED.
      changer->set_font_strikethrough( ip_font_strikethrough ).
    ENDIF.
    IF ip_font_underline IS SUPPLIED.
      changer->set_font_underline( ip_font_underline ).
    ENDIF.
    IF ip_font_underline_mode IS SUPPLIED.
      changer->set_font_underline_mode( ip_font_underline_mode ).
    ENDIF.

    IF ip_fill_filltype IS SUPPLIED.
      changer->set_fill_filltype( ip_fill_filltype ).
    ENDIF.
    IF ip_fill_rotation IS SUPPLIED.
      changer->set_fill_rotation( ip_fill_rotation ).
    ENDIF.
    IF ip_fill_fgcolor IS SUPPLIED.
      changer->set_fill_fgcolor( ip_fill_fgcolor ).
    ENDIF.
    IF ip_fill_fgcolor_rgb IS SUPPLIED.
      changer->set_fill_fgcolor_rgb( ip_fill_fgcolor_rgb ).
    ENDIF.
    IF ip_fill_fgcolor_indexed IS SUPPLIED.
      changer->set_fill_fgcolor_indexed( ip_fill_fgcolor_indexed ).
    ENDIF.
    IF ip_fill_fgcolor_theme IS SUPPLIED.
      changer->set_fill_fgcolor_theme( ip_fill_fgcolor_theme ).
    ENDIF.
    IF ip_fill_fgcolor_tint IS SUPPLIED.
      changer->set_fill_fgcolor_tint( ip_fill_fgcolor_tint ).
    ENDIF.

    IF ip_fill_bgcolor IS SUPPLIED.
      changer->set_fill_bgcolor( ip_fill_bgcolor ).
    ENDIF.
    IF ip_fill_bgcolor_rgb IS SUPPLIED.
      changer->set_fill_bgcolor_rgb( ip_fill_bgcolor_rgb ).
    ENDIF.
    IF ip_fill_bgcolor_indexed IS SUPPLIED.
      changer->set_fill_bgcolor_indexed( ip_fill_bgcolor_indexed ).
    ENDIF.
    IF ip_fill_bgcolor_theme IS SUPPLIED.
      changer->set_fill_bgcolor_theme( ip_fill_bgcolor_theme ).
    ENDIF.
    IF ip_fill_bgcolor_tint IS SUPPLIED.
      changer->set_fill_bgcolor_tint( ip_fill_bgcolor_tint ).
    ENDIF.

    IF ip_fill_gradtype_type IS SUPPLIED.
      changer->set_fill_gradtype_type( ip_fill_gradtype_type ).
    ENDIF.
    IF ip_fill_gradtype_degree IS SUPPLIED.
      changer->set_fill_gradtype_degree( ip_fill_gradtype_degree ).
    ENDIF.
    IF ip_fill_gradtype_bottom IS SUPPLIED.
      changer->set_fill_gradtype_bottom( ip_fill_gradtype_bottom ).
    ENDIF.
    IF ip_fill_gradtype_left IS SUPPLIED.
      changer->set_fill_gradtype_left( ip_fill_gradtype_left ).
    ENDIF.
    IF ip_fill_gradtype_top IS SUPPLIED.
      changer->set_fill_gradtype_top( ip_fill_gradtype_top ).
    ENDIF.
    IF ip_fill_gradtype_right IS SUPPLIED.
      changer->set_fill_gradtype_right( ip_fill_gradtype_right ).
    ENDIF.
    IF ip_fill_gradtype_position1 IS SUPPLIED.
      changer->set_fill_gradtype_position1( ip_fill_gradtype_position1 ).
    ENDIF.
    IF ip_fill_gradtype_position2 IS SUPPLIED.
      changer->set_fill_gradtype_position2( ip_fill_gradtype_position2 ).
    ENDIF.
    IF ip_fill_gradtype_position3 IS SUPPLIED.
      changer->set_fill_gradtype_position3( ip_fill_gradtype_position3 ).
    ENDIF.



    IF ip_borders_diagonal_mode IS SUPPLIED.
      changer->set_borders_diagonal_mode( ip_borders_diagonal_mode ).
    ENDIF.
    IF ip_alignment_horizontal IS SUPPLIED.
      changer->set_alignment_horizontal( ip_alignment_horizontal ).
    ENDIF.
    IF ip_alignment_vertical IS SUPPLIED.
      changer->set_alignment_vertical( ip_alignment_vertical ).
    ENDIF.
    IF ip_alignment_textrotation IS SUPPLIED.
      changer->set_alignment_textrotation( ip_alignment_textrotation ).
    ENDIF.
    IF ip_alignment_wraptext IS SUPPLIED.
      changer->set_alignment_wraptext( ip_alignment_wraptext ).
    ENDIF.
    IF ip_alignment_shrinktofit IS SUPPLIED.
      changer->set_alignment_shrinktofit( ip_alignment_shrinktofit ).
    ENDIF.
    IF ip_alignment_indent IS SUPPLIED.
      changer->set_alignment_indent( ip_alignment_indent ).
    ENDIF.
    IF ip_protection_hidden IS SUPPLIED.
      changer->set_protection_hidden( ip_protection_hidden ).
    ENDIF.
    IF ip_protection_locked IS SUPPLIED.
      changer->set_protection_locked( ip_protection_locked ).
    ENDIF.

    IF ip_borders_allborders_style IS SUPPLIED.
      changer->set_borders_allborders_style( ip_borders_allborders_style ).
    ENDIF.
    IF ip_borders_allborders_color IS SUPPLIED.
      changer->set_borders_allborders_color( ip_borders_allborders_color ).
    ENDIF.
    IF ip_borders_allbo_color_rgb IS SUPPLIED.
      changer->set_borders_allbo_color_rgb( ip_borders_allbo_color_rgb ).
    ENDIF.
    IF ip_borders_allbo_color_indexed IS SUPPLIED.
      changer->set_borders_allbo_color_indexe( ip_borders_allbo_color_indexed ).
    ENDIF.
    IF ip_borders_allbo_color_theme IS SUPPLIED.
      changer->set_borders_allbo_color_theme( ip_borders_allbo_color_theme ).
    ENDIF.
    IF ip_borders_allbo_color_tint IS SUPPLIED.
      changer->set_borders_allbo_color_tint( ip_borders_allbo_color_tint ).
    ENDIF.

    IF ip_borders_diagonal_style IS SUPPLIED.
      changer->set_borders_diagonal_style( ip_borders_diagonal_style ).
    ENDIF.
    IF ip_borders_diagonal_color IS SUPPLIED.
      changer->set_borders_diagonal_color( ip_borders_diagonal_color ).
    ENDIF.
    IF ip_borders_diagonal_color_rgb IS SUPPLIED.
      changer->set_borders_diagonal_color_rgb( ip_borders_diagonal_color_rgb ).
    ENDIF.
    IF ip_borders_diagonal_color_inde IS SUPPLIED.
      changer->set_borders_diagonal_color_ind( ip_borders_diagonal_color_inde ).
    ENDIF.
    IF ip_borders_diagonal_color_them IS SUPPLIED.
      changer->set_borders_diagonal_color_the( ip_borders_diagonal_color_them ).
    ENDIF.
    IF ip_borders_diagonal_color_tint IS SUPPLIED.
      changer->set_borders_diagonal_color_tin( ip_borders_diagonal_color_tint ).
    ENDIF.

    IF ip_borders_down_style IS SUPPLIED.
      changer->set_borders_down_style( ip_borders_down_style ).
    ENDIF.
    IF ip_borders_down_color IS SUPPLIED.
      changer->set_borders_down_color( ip_borders_down_color ).
    ENDIF.
    IF ip_borders_down_color_rgb IS SUPPLIED.
      changer->set_borders_down_color_rgb( ip_borders_down_color_rgb ).
    ENDIF.
    IF ip_borders_down_color_indexed IS SUPPLIED.
      changer->set_borders_down_color_indexed( ip_borders_down_color_indexed ).
    ENDIF.
    IF ip_borders_down_color_theme IS SUPPLIED.
      changer->set_borders_down_color_theme( ip_borders_down_color_theme ).
    ENDIF.
    IF ip_borders_down_color_tint IS SUPPLIED.
      changer->set_borders_down_color_tint( ip_borders_down_color_tint ).
    ENDIF.

    IF ip_borders_left_style IS SUPPLIED.
      changer->set_borders_left_style( ip_borders_left_style ).
    ENDIF.
    IF ip_borders_left_color IS SUPPLIED.
      changer->set_borders_left_color( ip_borders_left_color ).
    ENDIF.
    IF ip_borders_left_color_rgb IS SUPPLIED.
      changer->set_borders_left_color_rgb( ip_borders_left_color_rgb ).
    ENDIF.
    IF ip_borders_left_color_indexed IS SUPPLIED.
      changer->set_borders_left_color_indexed( ip_borders_left_color_indexed ).
    ENDIF.
    IF ip_borders_left_color_theme IS SUPPLIED.
      changer->set_borders_left_color_theme( ip_borders_left_color_theme ).
    ENDIF.
    IF ip_borders_left_color_tint IS SUPPLIED.
      changer->set_borders_left_color_tint( ip_borders_left_color_tint ).
    ENDIF.

    IF ip_borders_right_style IS SUPPLIED.
      changer->set_borders_right_style( ip_borders_right_style ).
    ENDIF.
    IF ip_borders_right_color IS SUPPLIED.
      changer->set_borders_right_color( ip_borders_right_color ).
    ENDIF.
    IF ip_borders_right_color_rgb IS SUPPLIED.
      changer->set_borders_right_color_rgb( ip_borders_right_color_rgb ).
    ENDIF.
    IF ip_borders_right_color_indexed IS SUPPLIED.
      changer->set_borders_right_color_indexe( ip_borders_right_color_indexed ).
    ENDIF.
    IF ip_borders_right_color_theme IS SUPPLIED.
      changer->set_borders_right_color_theme( ip_borders_right_color_theme ).
    ENDIF.
    IF ip_borders_right_color_tint IS SUPPLIED.
      changer->set_borders_right_color_tint( ip_borders_right_color_tint ).
    ENDIF.

    IF ip_borders_top_style IS SUPPLIED.
      changer->set_borders_top_style( ip_borders_top_style ).
    ENDIF.
    IF ip_borders_top_color IS SUPPLIED.
      changer->set_borders_top_color( ip_borders_top_color ).
    ENDIF.
    IF ip_borders_top_color_rgb IS SUPPLIED.
      changer->set_borders_top_color_rgb( ip_borders_top_color_rgb ).
    ENDIF.
    IF ip_borders_top_color_indexed IS SUPPLIED.
      changer->set_borders_top_color_indexed( ip_borders_top_color_indexed ).
    ENDIF.
    IF ip_borders_top_color_theme IS SUPPLIED.
      changer->set_borders_top_color_theme( ip_borders_top_color_theme ).
    ENDIF.
    IF ip_borders_top_color_tint IS SUPPLIED.
      changer->set_borders_top_color_tint( ip_borders_top_color_tint ).
    ENDIF.


    ep_guid = changer->apply( ip_worksheet = me
                              ip_column    = column
                              ip_row       = row ).


  ENDMETHOD.                    "CHANGE_CELL_STYLE


  METHOD check_cell_column_formula.

    FIELD-SYMBOLS <fs_column_formula> TYPE zcl_excel_worksheet=>mty_s_column_formula.

    IF ip_value IS NOT INITIAL OR ip_formula IS NOT INITIAL.
      zcx_excel=>raise_text( c_messages-formula_id_only_is_possible ).
    ENDIF.
    READ TABLE it_column_formulas WITH TABLE KEY id = ip_column_formula_id ASSIGNING <fs_column_formula>.
    IF sy-subrc <> 0.
      zcx_excel=>raise_text( c_messages-column_formula_id_not_found ).
    ENDIF.
    IF ip_row < <fs_column_formula>-table_top_left_row + 1
          OR ip_row > <fs_column_formula>-table_bottom_right_row + 1
          OR ip_column < <fs_column_formula>-table_left_column_int
          OR ip_column > <fs_column_formula>-table_right_column_int.
      zcx_excel=>raise_text( c_messages-formula_not_in_this_table ).
    ENDIF.
    IF ip_column <> <fs_column_formula>-column.
      zcx_excel=>raise_text( c_messages-formula_in_other_column ).
    ENDIF.

  ENDMETHOD.


  METHOD check_rtf.

    DATA: lo_style           TYPE REF TO zcl_excel_style,
          lo_iterator        TYPE REF TO zcl_excel_collection_iterator,
          lv_next_rtf_offset TYPE i,
          lv_tabix           TYPE i,
          lv_value           TYPE string,
          lv_val_length      TYPE i,
          ls_rtf             LIKE LINE OF ct_rtf.
    FIELD-SYMBOLS: <rtf> LIKE LINE OF ct_rtf.

    IF ip_style IS NOT SUPPLIED.
      ip_style = excel->get_default_style( ).
    ENDIF.

    lo_iterator = excel->get_styles_iterator( ).
    WHILE lo_iterator->has_next( ) = abap_true.
      lo_style ?= lo_iterator->get_next( ).
      IF lo_style->get_guid( ) = ip_style.
        EXIT.
      ENDIF.
      CLEAR lo_style.
    ENDWHILE.

    lv_next_rtf_offset = 0.
    LOOP AT ct_rtf ASSIGNING <rtf>.
      lv_tabix = sy-tabix.
      IF lv_next_rtf_offset < <rtf>-offset.
        ls_rtf-offset = lv_next_rtf_offset.
        ls_rtf-length = <rtf>-offset - lv_next_rtf_offset.
        ls_rtf-font   = lo_style->font->get_structure( ).
        INSERT ls_rtf INTO ct_rtf INDEX lv_tabix.
      ELSEIF lv_next_rtf_offset > <rtf>-offset.
        RAISE EXCEPTION TYPE zcx_excel
          EXPORTING
            error = 'Gaps or overlaps in RTF data offset/length specs'.
      ENDIF.
      lv_next_rtf_offset = <rtf>-offset + <rtf>-length.
    ENDLOOP.

    lv_value = ip_value.
    lv_val_length = strlen( lv_value ).
    IF lv_val_length > lv_next_rtf_offset.
      ls_rtf-offset = lv_next_rtf_offset.
      ls_rtf-length = lv_val_length - lv_next_rtf_offset.
      ls_rtf-font   = lo_style->font->get_structure( ).
      INSERT ls_rtf INTO TABLE ct_rtf.
    ELSEIF lv_val_length > lv_next_rtf_offset.
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'RTF specs length is not equal to value length'.
    ENDIF.

  ENDMETHOD.


  METHOD check_table_overlapping.

    DATA: lv_errormessage TYPE string,
          lv_column_int   TYPE zexcel_cell_column,
          lv_maxcol       TYPE i.
    FIELD-SYMBOLS:
          <ls_table_settings> TYPE zexcel_s_table_settings.

    lv_column_int = zcl_excel_common=>convert_column2int( is_table_settings-top_left_column ).
    lv_maxcol = zcl_excel_common=>convert_column2int( is_table_settings-bottom_right_column ).

    LOOP AT it_other_table_settings ASSIGNING <ls_table_settings>.

      IF  (    (  is_table_settings-top_left_row     GE <ls_table_settings>-top_left_row
              AND is_table_settings-top_left_row     LE <ls_table_settings>-bottom_right_row )
            OR
               (  is_table_settings-bottom_right_row GE <ls_table_settings>-top_left_row
              AND is_table_settings-bottom_right_row LE <ls_table_settings>-bottom_right_row )
          )
        AND
          (    (  lv_column_int GE zcl_excel_common=>convert_column2int( <ls_table_settings>-top_left_column )
              AND lv_column_int LE zcl_excel_common=>convert_column2int( <ls_table_settings>-bottom_right_column ) )
            OR
               (  lv_maxcol     GE zcl_excel_common=>convert_column2int( <ls_table_settings>-top_left_column )
              AND lv_maxcol     LE zcl_excel_common=>convert_column2int( <ls_table_settings>-bottom_right_column ) )
          ).
        lv_errormessage = 'Table overlaps with previously bound table and will not be added to worksheet.'(400).
        zcx_excel=>raise_text( lv_errormessage ).
      ENDIF.

    ENDLOOP.

  ENDMETHOD.


  METHOD class_constructor.

    c_messages-formula_id_only_is_possible = |{ 'If Formula ID is used, value and formula must be empty'(008) }|.
    c_messages-column_formula_id_not_found = |{ 'The Column Formula does not exist'(009) }|.
    c_messages-formula_not_in_this_table = |{ 'The cell uses a Column Formula which should be part of the same table'(010) }|.
    c_messages-formula_in_other_column = |{ 'The cell uses a Column Formula which is in a different column'(011) }|.

  ENDMETHOD.


  METHOD clear_initial_colorxfields.

    IF is_color-rgb IS INITIAL.
      CLEAR cs_xcolor-rgb.
    ENDIF.
    IF is_color-indexed IS INITIAL.
      CLEAR cs_xcolor-indexed.
    ENDIF.
    IF is_color-theme IS INITIAL.
      CLEAR cs_xcolor-theme.
    ENDIF.
    IF is_color-tint IS INITIAL.
      CLEAR cs_xcolor-tint.
    ENDIF.

  ENDMETHOD.


  METHOD constructor.
    DATA: lv_title TYPE zexcel_sheet_title.

    me->excel = ip_excel.

    me->guid = zcl_excel_obsolete_func_wrap=>guid_create( ).        " ins issue #379 - replacement for outdated function call

    IF ip_title IS NOT INITIAL.
      lv_title = ip_title.
    ELSE.
      lv_title = me->generate_title( ). " ins issue #154 - Names of worksheets
    ENDIF.

    me->set_title( ip_title = lv_title ).

    CREATE OBJECT sheet_setup.
    CREATE OBJECT styles_cond.
    CREATE OBJECT data_validations.
    CREATE OBJECT tables.
    CREATE OBJECT columns.
    CREATE OBJECT rows.
    CREATE OBJECT ranges. " issue #163
    CREATE OBJECT mo_pagebreaks.
    CREATE OBJECT drawings
      EXPORTING
        ip_type = zcl_excel_drawing=>type_image.
    CREATE OBJECT charts
      EXPORTING
        ip_type = zcl_excel_drawing=>type_chart.
    me->zif_excel_sheet_protection~initialize( ).
    me->zif_excel_sheet_properties~initialize( ).
    CREATE OBJECT hyperlinks.
    CREATE OBJECT comments. " (+) Issue #180

* initialize active cell coordinates
    active_cell-cell_row = 1.
    active_cell-cell_column = 1.

* inizialize dimension range
    lower_cell-cell_row     = 1.
    lower_cell-cell_column  = 1.
    upper_cell-cell_row     = 1.
    upper_cell-cell_column  = 1.

  ENDMETHOD.                    "CONSTRUCTOR


  METHOD convert_to_table.

    TYPES:
      BEGIN OF ts_field_conv,
        fieldname TYPE x031l-fieldname,
        convexit  TYPE x031l-convexit,
      END OF ts_field_conv,
      BEGIN OF ts_style_conv,
        cell_style TYPE zexcel_s_cell_data-cell_style,
        abap_type  TYPE abap_typekind,
      END OF ts_style_conv.

    DATA:
      lv_row_int          TYPE zexcel_cell_row,
      lv_column_int       TYPE zexcel_cell_column,
      lv_column_alpha     TYPE zexcel_cell_column_alpha,
      lt_field_catalog    TYPE zexcel_t_fieldcatalog,
      ls_field_catalog    TYPE zexcel_s_fieldcatalog,
      lv_value            TYPE string,
      lv_maxcol           TYPE i,
      lv_maxrow           TYPE i,
      lt_field_conv       TYPE TABLE OF ts_field_conv,
      lt_comp             TYPE abap_component_tab,
      ls_comp             TYPE abap_componentdescr,
      lo_line_type        TYPE REF TO cl_abap_structdescr,
      lo_tab_type         TYPE REF TO cl_abap_tabledescr,
      lr_data             TYPE REF TO data,
      lt_comp_view        TYPE abap_component_view_tab,
      ls_comp_view        TYPE abap_simple_componentdescr,
      lt_ddic_object      TYPE dd_x031l_table,
      lt_ddic_object_comp TYPE dd_x031l_table,
      ls_ddic_object      TYPE x031l,
      lt_style_conv       TYPE TABLE OF ts_style_conv,
      ls_style_conv       TYPE ts_style_conv,
      ls_stylemapping     TYPE zexcel_s_stylemapping,
      lv_format_code      TYPE zexcel_number_format,
      lv_float            TYPE f,
      lt_map_excel_row    TYPE TABLE OF i,
      lv_index            TYPE i,
      lv_index_col        TYPE i.

    FIELD-SYMBOLS:
      <lt_data>          TYPE STANDARD TABLE,
      <ls_data>          TYPE data,
      <lv_data>          TYPE data,
      <lt_data2>         TYPE STANDARD TABLE,
      <ls_data2>         TYPE data,
      <lv_data2>         TYPE data,
      <ls_field_conv>    TYPE ts_field_conv,
      <ls_ddic_object>   TYPE x031l,
      <ls_sheet_content> TYPE zexcel_s_cell_data,
      <fs_typekind_int8> TYPE abap_typekind.

    CLEAR: et_data, er_data.

    lv_maxcol = get_highest_column( ).
    lv_maxrow = get_highest_row( ).


    " Field catalog
    lt_field_catalog = it_field_catalog.
    IF lt_field_catalog IS INITIAL.
      IF et_data IS SUPPLIED.
        lt_field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = et_data ).
      ELSE.
        DO lv_maxcol TIMES.
          ls_field_catalog-position = sy-index.
          ls_field_catalog-fieldname = 'COL_' && sy-index.
          ls_field_catalog-dynpfld = abap_true.
          APPEND ls_field_catalog TO lt_field_catalog.
        ENDDO.
      ENDIF.
    ENDIF.

    SORT lt_field_catalog BY position.
    DELETE lt_field_catalog WHERE dynpfld NE abap_true.
    CHECK: lt_field_catalog IS NOT INITIAL.


    " Create dynamic table string columns
    ls_comp-type = cl_abap_elemdescr=>get_string( ).
    LOOP AT lt_field_catalog INTO ls_field_catalog.
      ls_comp-name = ls_field_catalog-fieldname.
      APPEND ls_comp TO lt_comp.
    ENDLOOP.
    lo_line_type = cl_abap_structdescr=>create( lt_comp ).
    lo_tab_type = cl_abap_tabledescr=>create( lo_line_type ).
    CREATE DATA er_data TYPE HANDLE lo_tab_type.
    ASSIGN er_data->* TO <lt_data>.


    " Collect field conversion rules
    IF et_data IS SUPPLIED.
*      lt_ddic_object = get_ddic_object( et_data ).
      lo_tab_type ?= cl_abap_tabledescr=>describe_by_data( et_data ).
      lo_line_type ?= lo_tab_type->get_table_line_type( ).
      lo_line_type->get_ddic_object(
        RECEIVING
          p_object = lt_ddic_object
        EXCEPTIONS
          OTHERS   = 3
      ).
      IF lt_ddic_object IS INITIAL.
        lt_comp_view = lo_line_type->get_included_view( ).
        LOOP AT lt_comp_view INTO ls_comp_view.
          ls_comp_view-type->get_ddic_object(
            RECEIVING
              p_object = lt_ddic_object_comp
            EXCEPTIONS
              OTHERS   = 3
          ).
          IF lt_ddic_object_comp IS NOT INITIAL.
            READ TABLE lt_ddic_object_comp INTO ls_ddic_object INDEX 1.
            ls_ddic_object-fieldname = ls_comp_view-name.
            APPEND ls_ddic_object TO lt_ddic_object.
          ENDIF.
        ENDLOOP.
      ENDIF.

      SORT lt_ddic_object BY fieldname.
      LOOP AT lt_field_catalog INTO ls_field_catalog.
        APPEND INITIAL LINE TO lt_field_conv ASSIGNING <ls_field_conv>.
        MOVE-CORRESPONDING ls_field_catalog TO <ls_field_conv>.
        READ TABLE lt_ddic_object ASSIGNING <ls_ddic_object> WITH KEY fieldname = <ls_field_conv>-fieldname BINARY SEARCH.
        CHECK: sy-subrc EQ 0.

        ASSIGN ('CL_ABAP_TYPEDESCR=>TYPEKIND_INT8') TO <fs_typekind_int8>.
        IF sy-subrc <> 0.
          ASSIGN space TO <fs_typekind_int8>. "not used as typekind!
        ENDIF.

        CASE <ls_ddic_object>-exid.
          WHEN cl_abap_typedescr=>typekind_int
            OR cl_abap_typedescr=>typekind_int1
            OR <fs_typekind_int8>
            OR cl_abap_typedescr=>typekind_int2
            OR cl_abap_typedescr=>typekind_packed
            OR cl_abap_typedescr=>typekind_decfloat
            OR cl_abap_typedescr=>typekind_decfloat16
            OR cl_abap_typedescr=>typekind_decfloat34
            OR cl_abap_typedescr=>typekind_float.
            " Numbers
            <ls_field_conv>-convexit = cl_abap_typedescr=>typekind_float.
          WHEN OTHERS.
            <ls_field_conv>-convexit = <ls_ddic_object>-convexit.
        ENDCASE.
      ENDLOOP.
    ENDIF.

    " Date & Time in excel style
    LOOP AT me->sheet_content ASSIGNING <ls_sheet_content> WHERE cell_style IS NOT INITIAL AND data_type IS INITIAL.
      ls_style_conv-cell_style = <ls_sheet_content>-cell_style.
      APPEND ls_style_conv TO lt_style_conv.
    ENDLOOP.
    IF lt_style_conv IS NOT INITIAL.
      SORT lt_style_conv BY cell_style.
      DELETE ADJACENT DUPLICATES FROM lt_style_conv COMPARING cell_style.

      LOOP AT lt_style_conv INTO ls_style_conv.

        ls_stylemapping = me->excel->get_style_to_guid( ls_style_conv-cell_style ).
        lv_format_code = ls_stylemapping-complete_style-number_format-format_code.
        " https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68
        IF lv_format_code CS ';'.
          lv_format_code = lv_format_code(sy-fdpos).
        ENDIF.
        CHECK: lv_format_code NA '#?'.

        " Remove color pattern
        REPLACE ALL OCCURRENCES OF REGEX '\[\L[^]]*\]' IN lv_format_code WITH ''.

        IF lv_format_code CA 'yd' OR lv_format_code EQ zcl_excel_style_number_format=>c_format_date_std.
          " DATE = yyyymmdd
          ls_style_conv-abap_type = cl_abap_typedescr=>typekind_date.
        ELSEIF lv_format_code CA 'hs'.
          " TIME = hhmmss
          ls_style_conv-abap_type = cl_abap_typedescr=>typekind_time.
        ELSE.
          DELETE lt_style_conv.
          CONTINUE.
        ENDIF.

        MODIFY lt_style_conv FROM ls_style_conv TRANSPORTING abap_type.

      ENDLOOP.
    ENDIF.


*--------------------------------------------------------------------*
* Start of convert content
*--------------------------------------------------------------------*
    READ TABLE me->sheet_content TRANSPORTING NO FIELDS WITH KEY cell_row = iv_begin_row.
    IF sy-subrc EQ 0.
      lv_index = sy-tabix.
    ENDIF.

    LOOP AT me->sheet_content ASSIGNING <ls_sheet_content> FROM lv_index.
      AT NEW cell_row.
        IF iv_end_row <> 0
        AND <ls_sheet_content>-cell_row > iv_end_row.
          EXIT.
        ENDIF.
        " New line
        APPEND INITIAL LINE TO <lt_data> ASSIGNING <ls_data>.
        lv_index = sy-tabix.
      ENDAT.

      IF <ls_sheet_content>-cell_value IS NOT INITIAL.
        ASSIGN COMPONENT <ls_sheet_content>-cell_column OF STRUCTURE <ls_data> TO <lv_data>.
        IF sy-subrc EQ 0.
          " value
          <lv_data> = <ls_sheet_content>-cell_value.

          " field conversion
          READ TABLE lt_field_conv ASSIGNING <ls_field_conv> INDEX <ls_sheet_content>-cell_column.
          IF sy-subrc EQ 0 AND <ls_field_conv>-convexit IS NOT INITIAL.
            CASE <ls_field_conv>-convexit.
              WHEN cl_abap_typedescr=>typekind_float.
                lv_float = zcl_excel_common=>excel_string_to_number( <ls_sheet_content>-cell_value ).
                <lv_data> = |{ lv_float NUMBER = RAW }|.
              WHEN 'ALPHA'.
                CALL FUNCTION 'CONVERSION_EXIT_ALPHA_OUTPUT'
                  EXPORTING
                    input  = <ls_sheet_content>-cell_value
                  IMPORTING
                    output = <lv_data>.
            ENDCASE.
          ENDIF.

          " style conversion
          IF <ls_sheet_content>-cell_style IS NOT INITIAL.
            READ TABLE lt_style_conv INTO ls_style_conv WITH KEY cell_style = <ls_sheet_content>-cell_style BINARY SEARCH.
            IF sy-subrc EQ 0.
              CASE ls_style_conv-abap_type.
                WHEN cl_abap_typedescr=>typekind_date.
                  <lv_data> = zcl_excel_common=>excel_string_to_date( ip_value = <ls_sheet_content>-cell_value
                                                                      ip_exact = abap_true ).
                WHEN cl_abap_typedescr=>typekind_time.
                  <lv_data> = zcl_excel_common=>excel_string_to_time( <ls_sheet_content>-cell_value ).
              ENDCASE.
            ENDIF.
          ENDIF.

          " condense
          CONDENSE <lv_data>.
        ENDIF.
      ENDIF.

      AT END OF cell_row.
        " Delete empty line
        IF <ls_data> IS INITIAL.
          DELETE <lt_data> INDEX lv_index.
        ELSE.
          APPEND <ls_sheet_content>-cell_row TO lt_map_excel_row.
        ENDIF.
      ENDAT.
    ENDLOOP.
*--------------------------------------------------------------------*
* End of convert content
*--------------------------------------------------------------------*


    IF et_data IS SUPPLIED.
*      MOVE-CORRESPONDING <lt_data> TO et_data.
      LOOP AT <lt_data> ASSIGNING <ls_data>.
        APPEND INITIAL LINE TO et_data ASSIGNING <ls_data2>.
        MOVE-CORRESPONDING <ls_data> TO <ls_data2>.
      ENDLOOP.
    ENDIF.

    " Apply conversion exit.
    LOOP AT lt_field_conv ASSIGNING <ls_field_conv>
     WHERE convexit = 'ALPHA'.
      LOOP AT et_data ASSIGNING <ls_data>.
        ASSIGN COMPONENT <ls_field_conv>-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
        CHECK: sy-subrc EQ 0 AND <lv_data> IS NOT INITIAL.
        CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
          EXPORTING
            input  = <lv_data>
          IMPORTING
            output = <lv_data>.
      ENDLOOP.
    ENDLOOP.

  ENDMETHOD.


  METHOD create_data_conv_exit_length.
    DATA: lo_addit    TYPE REF TO cl_abap_elemdescr,
          ls_dfies    TYPE dfies,
          l_function  TYPE funcname,
          l_value(50) TYPE c.

    lo_addit ?= cl_abap_typedescr=>describe_by_data( ip_value ).
    lo_addit->get_ddic_field( RECEIVING  p_flddescr   = ls_dfies
                              EXCEPTIONS not_found    = 1
                                         no_ddic_type = 2
                                         OTHERS       = 3 ) .
    IF sy-subrc = 0 AND ls_dfies-convexit IS NOT INITIAL.
      CREATE DATA ep_value TYPE c LENGTH ls_dfies-outputlen.
    ELSE.
      CREATE DATA ep_value LIKE ip_value.
    ENDIF.

  ENDMETHOD.


  METHOD delete_merge.

    DATA: lv_column TYPE i.
*--------------------------------------------------------------------*
* If cell information is passed delete merge including this cell,
* otherwise delete all merges
*--------------------------------------------------------------------*
    IF   ip_cell_column IS INITIAL
      OR ip_cell_row    IS INITIAL.
      CLEAR me->mt_merged_cells.
    ELSE.
      lv_column = zcl_excel_common=>convert_column2int( ip_cell_column ).

      LOOP AT me->mt_merged_cells TRANSPORTING NO FIELDS
      WHERE row_from <= ip_cell_row AND row_to >= ip_cell_row
        AND col_from <= lv_column AND col_to >= lv_column.
        DELETE me->mt_merged_cells.
        EXIT.
      ENDLOOP.
    ENDIF.

  ENDMETHOD.                    "DELETE_MERGE


  METHOD delete_row_outline.

    DELETE me->mt_row_outlines WHERE row_from = iv_row_from
                                 AND row_to   = iv_row_to.
    IF sy-subrc <> 0.  " didn't find outline that was to be deleted
      zcx_excel=>raise_text( 'Row outline to be deleted does not exist' ).
    ENDIF.

  ENDMETHOD.                    "DELETE_ROW_OUTLINE


  METHOD freeze_panes.

    IF ip_num_columns IS NOT SUPPLIED AND ip_num_rows IS NOT SUPPLIED.
      zcx_excel=>raise_text( 'Pleas provide number of rows and/or columns to freeze' ).
    ENDIF.

    IF ip_num_columns IS SUPPLIED AND ip_num_columns <= 0.
      zcx_excel=>raise_text( 'Number of columns to freeze should be positive' ).
    ENDIF.

    IF ip_num_rows IS SUPPLIED AND ip_num_rows <= 0.
      zcx_excel=>raise_text( 'Number of rows to freeze should be positive' ).
    ENDIF.

    freeze_pane_cell_column = ip_num_columns + 1.
    freeze_pane_cell_row = ip_num_rows + 1.
  ENDMETHOD.                    "FREEZE_PANES


  METHOD generate_title.
    DATA: lo_worksheets_iterator TYPE REF TO zcl_excel_collection_iterator,
          lo_worksheet           TYPE REF TO zcl_excel_worksheet.

    DATA: t_titles    TYPE HASHED TABLE OF zexcel_sheet_title WITH UNIQUE KEY table_line,
          title       TYPE zexcel_sheet_title,
          sheetnumber TYPE i.

* Get list of currently used titles
    lo_worksheets_iterator = me->excel->get_worksheets_iterator( ).
    WHILE lo_worksheets_iterator->has_next( ) = abap_true.
      lo_worksheet ?= lo_worksheets_iterator->get_next( ).
      title = lo_worksheet->get_title( ).
      INSERT title INTO TABLE t_titles.
      ADD 1 TO sheetnumber.
    ENDWHILE.

* Now build sheetnumber.  Increase counter until we hit a number that is not used so far
    ADD 1 TO sheetnumber.  " Start counting with next number
    DO.
      title = sheetnumber.
      SHIFT title LEFT DELETING LEADING space.
      CONCATENATE 'Sheet'(001) title INTO ep_title.
      INSERT ep_title INTO TABLE t_titles.
      IF sy-subrc = 0.  " Title not used so far --> take it
        EXIT.
      ENDIF.

      ADD 1 TO sheetnumber.
    ENDDO.
  ENDMETHOD.                    "GENERATE_TITLE


  METHOD get_active_cell.

    DATA: lv_active_column TYPE zexcel_cell_column_alpha,
          lv_active_row    TYPE string.

    lv_active_column = zcl_excel_common=>convert_column2alpha( active_cell-cell_column ).
    lv_active_row    = active_cell-cell_row.
    SHIFT lv_active_row RIGHT DELETING TRAILING space.
    SHIFT lv_active_row LEFT DELETING LEADING space.
    CONCATENATE lv_active_column lv_active_row INTO ep_active_cell.

  ENDMETHOD.                    "GET_ACTIVE_CELL


  METHOD get_cell.

    DATA: lv_column        TYPE zexcel_cell_column,
          lv_row           TYPE zexcel_cell_row,
          ls_sheet_content TYPE zexcel_s_cell_data.

    normalize_columnrow_parameter( EXPORTING ip_columnrow = ip_columnrow
                                             ip_column    = ip_column
                                             ip_row       = ip_row
                                   IMPORTING ep_column    = lv_column
                                             ep_row       = lv_row ).

    READ TABLE sheet_content INTO ls_sheet_content WITH TABLE KEY cell_row     = lv_row
                                                                  cell_column  = lv_column.

    ep_rc       = sy-subrc.
    ep_value    = ls_sheet_content-cell_value.
    ep_guid     = ls_sheet_content-cell_style.       " issue 139 - added this to be used for columnwidth calculation
    ep_formula  = ls_sheet_content-cell_formula.
    IF et_rtf IS SUPPLIED AND ls_sheet_content-rtf_tab IS NOT INITIAL.
      et_rtf = ls_sheet_content-rtf_tab.
    ENDIF.

    " Addition to solve issue #120, contribution by Stefan Schmöcker
    DATA: style_iterator TYPE REF TO zcl_excel_collection_iterator,
          style          TYPE REF TO zcl_excel_style.
    IF ep_style IS SUPPLIED.
      CLEAR ep_style.
      style_iterator = me->excel->get_styles_iterator( ).
      WHILE style_iterator->has_next( ) = abap_true.
        style ?= style_iterator->get_next( ).
        IF style->get_guid( ) = ls_sheet_content-cell_style.
          ep_style = style.
          EXIT.
        ENDIF.
      ENDWHILE.
    ENDIF.
  ENDMETHOD.                    "GET_CELL


  METHOD get_column.

    DATA: lv_column TYPE zexcel_cell_column.

    lv_column = zcl_excel_common=>convert_column2int( ip_column ).

    eo_column = me->columns->get( ip_index = lv_column ).

    IF eo_column IS NOT BOUND.
      eo_column = me->add_new_column( ip_column ).
    ENDIF.

  ENDMETHOD.                    "GET_COLUMN


  METHOD get_columns.

    DATA: columns TYPE TABLE OF i,
          column  TYPE i.
    FIELD-SYMBOLS:
          <sheet_cell> TYPE zexcel_s_cell_data.

    LOOP AT sheet_content ASSIGNING <sheet_cell>.
      COLLECT <sheet_cell>-cell_column INTO columns.
    ENDLOOP.

    LOOP AT columns INTO column.
      " This will create the column instance if it doesn't exist
      get_column( column ).
    ENDLOOP.

    eo_columns = me->columns.
  ENDMETHOD.                    "GET_COLUMNS


  METHOD get_columns_iterator.

    get_columns( ).
    eo_iterator = me->columns->get_iterator( ).

  ENDMETHOD.                    "GET_COLUMNS_ITERATOR


  METHOD get_comments.
    DATA: lo_comment  TYPE REF TO zcl_excel_comment,
          lo_iterator TYPE REF TO zcl_excel_collection_iterator.

    CREATE OBJECT r_comments.

    lo_iterator = comments->get_iterator( ).
    WHILE lo_iterator->has_next( ) = abap_true.
      lo_comment ?= lo_iterator->get_next( ).
      r_comments->include( lo_comment ).
    ENDWHILE.

  ENDMETHOD.                    "get_comments


  METHOD get_comments_iterator.
    eo_iterator = comments->get_iterator( ).

  ENDMETHOD.                    "get_comments_iterator


  METHOD get_data_validations_iterator.

    eo_iterator = me->data_validations->get_iterator( ).
  ENDMETHOD.                    "GET_DATA_VALIDATIONS_ITERATOR


  METHOD get_data_validations_size.
    ep_size = me->data_validations->size( ).
  ENDMETHOD.                    "GET_DATA_VALIDATIONS_SIZE


  METHOD get_default_column.
    IF me->column_default IS NOT BOUND.
      CREATE OBJECT me->column_default
        EXPORTING
          ip_index     = 'A'         " ????
          ip_worksheet = me
          ip_excel     = me->excel.
    ENDIF.

    eo_column = me->column_default.
  ENDMETHOD.                    "GET_DEFAULT_COLUMN


  METHOD get_default_excel_date_format.
    CONSTANTS: c_lang_e TYPE lang VALUE 'E'.

    IF default_excel_date_format IS NOT INITIAL.
      ep_default_excel_date_format = default_excel_date_format.
      RETURN.
    ENDIF.

    "try to get defaults
    TRY.
        cl_abap_datfm=>get_date_format_des( EXPORTING im_langu = c_lang_e
                                            IMPORTING ex_dateformat = default_excel_date_format ).
      CATCH cx_abap_datfm_format_unknown.

    ENDTRY.

    " and fallback to fixed format
    IF default_excel_date_format IS INITIAL.
      default_excel_date_format = zcl_excel_style_number_format=>c_format_date_ddmmyyyydot.
    ENDIF.

    ep_default_excel_date_format = default_excel_date_format.
  ENDMETHOD.                    "GET_DEFAULT_EXCEL_DATE_FORMAT


  METHOD get_default_excel_time_format.
    DATA: l_timefm TYPE xutimefm.

    IF default_excel_time_format IS NOT INITIAL.
      ep_default_excel_time_format = default_excel_time_format.
      RETURN.
    ENDIF.

* Let's get default
    l_timefm = cl_abap_timefm=>get_environment_timefm( ).
    CASE l_timefm.
      WHEN 0.
*0  24 Hour Format (Example: 12:05:10)
        default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time6.
      WHEN 1.
*1  12 Hour Format (Example: 12:05:10 PM)
        default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time2.
      WHEN 2.
*2  12 Hour Format (Example: 12:05:10 pm) for now all the same. no chnage upper lower
        default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time2.
      WHEN 3.
*3  Hours from 0 to 11 (Example: 00:05:10 PM)  for now all the same. no chnage upper lower
        default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time2.
      WHEN 4.
*4  Hours from 0 to 11 (Example: 00:05:10 pm)  for now all the same. no chnage upper lower
        default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time2.
      WHEN OTHERS.
        " and fallback to fixed format
        default_excel_time_format = zcl_excel_style_number_format=>c_format_date_time6.
    ENDCASE.

    ep_default_excel_time_format = default_excel_time_format.
  ENDMETHOD.                    "GET_DEFAULT_EXCEL_TIME_FORMAT


  METHOD get_default_row.
    IF me->row_default IS NOT BOUND.
      CREATE OBJECT me->row_default.
    ENDIF.

    eo_row = me->row_default.
  ENDMETHOD.                    "GET_DEFAULT_ROW


  METHOD get_dimension_range.

    me->update_dimension_range( ).
    IF upper_cell EQ lower_cell. "only one cell
      " Worksheet not filled
      IF upper_cell-cell_coords IS INITIAL.
        ep_dimension_range = 'A1'.
      ELSE.
        ep_dimension_range = upper_cell-cell_coords.
      ENDIF.
    ELSE.
      CONCATENATE upper_cell-cell_coords ':' lower_cell-cell_coords INTO ep_dimension_range.
    ENDIF.

  ENDMETHOD.                    "GET_DIMENSION_RANGE


  METHOD get_drawings.

    DATA: lo_drawing  TYPE REF TO zcl_excel_drawing,
          lo_iterator TYPE REF TO zcl_excel_collection_iterator.

    CASE ip_type.
      WHEN zcl_excel_drawing=>type_image.
        r_drawings = drawings.
      WHEN zcl_excel_drawing=>type_chart.
        r_drawings = charts.
      WHEN space.
        CREATE OBJECT r_drawings
          EXPORTING
            ip_type = ''.

        lo_iterator = drawings->get_iterator( ).
        WHILE lo_iterator->has_next( ) = abap_true.
          lo_drawing ?= lo_iterator->get_next( ).
          r_drawings->include( lo_drawing ).
        ENDWHILE.
        lo_iterator = charts->get_iterator( ).
        WHILE lo_iterator->has_next( ) = abap_true.
          lo_drawing ?= lo_iterator->get_next( ).
          r_drawings->include( lo_drawing ).
        ENDWHILE.
      WHEN OTHERS.
    ENDCASE.
  ENDMETHOD.                    "GET_DRAWINGS


  METHOD get_drawings_iterator.
    CASE ip_type.
      WHEN zcl_excel_drawing=>type_image.
        eo_iterator = drawings->get_iterator( ).
      WHEN zcl_excel_drawing=>type_chart.
        eo_iterator = charts->get_iterator( ).
    ENDCASE.
  ENDMETHOD.                    "GET_DRAWINGS_ITERATOR


  METHOD get_freeze_cell.
    ep_row = me->freeze_pane_cell_row.
    ep_column = me->freeze_pane_cell_column.
  ENDMETHOD.                    "GET_FREEZE_CELL


  METHOD get_guid.

    ep_guid = me->guid.

  ENDMETHOD.                    "GET_GUID


  METHOD get_header_footer_drawings.
    DATA: ls_odd_header  TYPE zexcel_s_worksheet_head_foot,
          ls_odd_footer  TYPE zexcel_s_worksheet_head_foot,
          ls_even_header TYPE zexcel_s_worksheet_head_foot,
          ls_even_footer TYPE zexcel_s_worksheet_head_foot,
          ls_hd_ft       TYPE zexcel_s_worksheet_head_foot.

    FIELD-SYMBOLS: <fs_drawings> TYPE zexcel_s_drawings.

    me->sheet_setup->get_header_footer( IMPORTING ep_odd_header = ls_odd_header
                                                  ep_odd_footer = ls_odd_footer
                                                  ep_even_header = ls_even_header
                                                  ep_even_footer = ls_even_footer ).

**********************************************************************
*** Odd header
    ls_hd_ft = ls_odd_header.
    IF ls_hd_ft-left_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-left_image.
    ENDIF.
    IF ls_hd_ft-right_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-right_image.
    ENDIF.
    IF ls_hd_ft-center_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-center_image.
    ENDIF.

**********************************************************************
*** Odd footer
    ls_hd_ft = ls_odd_footer.
    IF ls_hd_ft-left_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-left_image.
    ENDIF.
    IF ls_hd_ft-right_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-right_image.
    ENDIF.
    IF ls_hd_ft-center_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-center_image.
    ENDIF.

**********************************************************************
*** Even header
    ls_hd_ft = ls_even_header.
    IF ls_hd_ft-left_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-left_image.
    ENDIF.
    IF ls_hd_ft-right_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-right_image.
    ENDIF.
    IF ls_hd_ft-center_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-center_image.
    ENDIF.

**********************************************************************
*** Even footer
    ls_hd_ft = ls_even_footer.
    IF ls_hd_ft-left_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-left_image.
    ENDIF.
    IF ls_hd_ft-right_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-right_image.
    ENDIF.
    IF ls_hd_ft-center_image IS NOT INITIAL.
      APPEND INITIAL LINE TO rt_drawings ASSIGNING <fs_drawings>.
      <fs_drawings>-drawing = ls_hd_ft-center_image.
    ENDIF.

  ENDMETHOD.                    "get_header_footer_drawings


  METHOD get_highest_column.
    me->update_dimension_range( ).
    r_highest_column = me->lower_cell-cell_column.
  ENDMETHOD.                    "GET_HIGHEST_COLUMN


  METHOD get_highest_row.
    me->update_dimension_range( ).
    r_highest_row = me->lower_cell-cell_row.
  ENDMETHOD.                    "GET_HIGHEST_ROW


  METHOD get_hyperlinks_iterator.
    eo_iterator = hyperlinks->get_iterator( ).
  ENDMETHOD.                    "GET_HYPERLINKS_ITERATOR


  METHOD get_hyperlinks_size.
    ep_size = hyperlinks->size( ).
  ENDMETHOD.                    "GET_HYPERLINKS_SIZE


  METHOD get_ignored_errors.
    rt_ignored_errors = mt_ignored_errors.
  ENDMETHOD.


  METHOD get_merge.

    FIELD-SYMBOLS: <ls_merged_cell> LIKE LINE OF me->mt_merged_cells.

    DATA: lv_col_from    TYPE string,
          lv_col_to      TYPE string,
          lv_row_from    TYPE string,
          lv_row_to      TYPE string,
          lv_merge_range TYPE string.

    LOOP AT me->mt_merged_cells ASSIGNING <ls_merged_cell>.

      lv_col_from = zcl_excel_common=>convert_column2alpha( <ls_merged_cell>-col_from ).
      lv_col_to   = zcl_excel_common=>convert_column2alpha( <ls_merged_cell>-col_to   ).
      lv_row_from = <ls_merged_cell>-row_from.
      lv_row_to   = <ls_merged_cell>-row_to  .
      CONCATENATE lv_col_from lv_row_from ':' lv_col_to lv_row_to
         INTO lv_merge_range.
      CONDENSE lv_merge_range NO-GAPS.
      APPEND lv_merge_range TO merge_range.

    ENDLOOP.

  ENDMETHOD.                    "GET_MERGE


  METHOD get_pagebreaks.
    ro_pagebreaks = mo_pagebreaks.
  ENDMETHOD.                    "GET_PAGEBREAKS


  METHOD get_ranges_iterator.

    eo_iterator = me->ranges->get_iterator( ).

  ENDMETHOD.                    "GET_RANGES_ITERATOR


  METHOD get_row.
    eo_row = me->rows->get( ip_index = ip_row ).

    IF eo_row IS NOT BOUND.
      eo_row = me->add_new_row( ip_row ).
    ENDIF.
  ENDMETHOD.                    "GET_ROW


  METHOD get_rows.

    DATA: row TYPE i.
    FIELD-SYMBOLS: <sheet_cell> TYPE zexcel_s_cell_data.

    IF sheet_content IS NOT INITIAL.

      row = 0.
      DO.
        " Find the next row
        READ TABLE sheet_content ASSIGNING <sheet_cell> WITH KEY cell_row = row.
        CASE sy-subrc.
          WHEN 4.
            " row doesn't exist, but it exists another row, SY-TABIX points to the first cell in this row.
            READ TABLE sheet_content ASSIGNING <sheet_cell> INDEX sy-tabix.
            ASSERT sy-subrc = 0.
            row = <sheet_cell>-cell_row.
          WHEN 8.
            " it was the last available row
            EXIT.
        ENDCASE.
        " This will create the row instance if it doesn't exist
        get_row( row ).
        row = row + 1.
      ENDDO.

    ENDIF.

    eo_rows = me->rows.
  ENDMETHOD.                    "GET_ROWS


  METHOD get_rows_iterator.

    get_rows( ).
    eo_iterator = me->rows->get_iterator( ).

  ENDMETHOD.                    "GET_ROWS_ITERATOR


  METHOD get_row_outlines.

    rt_row_outlines = me->mt_row_outlines.

  ENDMETHOD.                    "GET_ROW_OUTLINES


  METHOD get_style_cond.

    DATA: lo_style_iterator TYPE REF TO zcl_excel_collection_iterator,
          lo_style_cond     TYPE REF TO zcl_excel_style_cond.

    lo_style_iterator = me->get_style_cond_iterator( ).
    WHILE lo_style_iterator->has_next( ) = abap_true.
      lo_style_cond ?= lo_style_iterator->get_next( ).
      IF lo_style_cond->get_guid( ) = ip_guid.
        eo_style_cond = lo_style_cond.
        EXIT.
      ENDIF.
    ENDWHILE.

  ENDMETHOD.                    "GET_STYLE_COND


  METHOD get_style_cond_iterator.

    eo_iterator = styles_cond->get_iterator( ).
  ENDMETHOD.                    "GET_STYLE_COND_ITERATOR


  METHOD get_tabcolor.
    ev_tabcolor = me->tabcolor.
  ENDMETHOD.                    "GET_TABCOLOR


  METHOD get_table.
*--------------------------------------------------------------------*
* Comment D. Rauchenstein
* With this method, we get a fully functional Excel Upload, which solves
* a few issues of the other excel upload tools
* ZBCABA_ALSM_EXCEL_UPLOAD_EXT: Reads only up to 50 signs per Cell, Limit
* in row-Numbers. Other have Limitations of Lines, or you are not able
* to ignore filters or choosing the right tab.
*
* To get a fully functional XLSX Upload, you can use it e.g. with method
* CL_EXCEL_READER_2007->ZIF_EXCEL_READER~LOAD_FILE()
*--------------------------------------------------------------------*

    FIELD-SYMBOLS: <ls_line> TYPE data.
    FIELD-SYMBOLS: <lv_value> TYPE data.

    DATA lv_actual_row TYPE int4.
    DATA lv_actual_row_string TYPE string.
    DATA lv_actual_col TYPE int4.
    DATA lv_actual_col_string TYPE string.
    DATA lv_errormessage TYPE string.
    DATA lv_max_col TYPE zexcel_cell_column.
    DATA lv_max_row TYPE int4.
    DATA lv_delta_col TYPE int4.
    DATA lv_value  TYPE zexcel_cell_value.
    DATA lv_rc  TYPE sysubrc.
    DATA lx_conversion_error TYPE REF TO cx_sy_conversion_error.
    DATA lv_float TYPE f.
    DATA lv_type.
    DATA lv_tabix TYPE i.

    lv_max_col =  me->get_highest_column( ).
    IF iv_max_col IS SUPPLIED AND iv_max_col < lv_max_col.
      lv_max_col = iv_max_col.
    ENDIF.
    lv_max_row =  me->get_highest_row( ).
    IF iv_max_row IS SUPPLIED AND iv_max_row < lv_max_row.
      lv_max_row = iv_max_row.
    ENDIF.

*--------------------------------------------------------------------*
* The row counter begins with 1 and should be corrected with the skips
*--------------------------------------------------------------------*
    lv_actual_row =  iv_skipped_rows + 1.
    lv_actual_col =  iv_skipped_cols + 1.


    TRY.
*--------------------------------------------------------------------*
* Check if we the basic features are possible with given "any table"
*--------------------------------------------------------------------*
        APPEND INITIAL LINE TO et_table ASSIGNING <ls_line>.
        IF sy-subrc <> 0 OR <ls_line> IS NOT ASSIGNED.

          lv_errormessage = 'Error at inserting new Line to internal Table'(002).
          zcx_excel=>raise_text( lv_errormessage ).

        ELSE.
          lv_delta_col = lv_max_col - iv_skipped_cols.
          ASSIGN COMPONENT lv_delta_col OF STRUCTURE <ls_line> TO <lv_value>.
          IF sy-subrc <> 0 OR <lv_value> IS NOT ASSIGNED.
            lv_errormessage = 'Internal table has less columns than excel'(003).
            zcx_excel=>raise_text( lv_errormessage ).
          ELSE.
*--------------------------------------------------------------------*
*now we are ready for handle the table data
*--------------------------------------------------------------------*
            CLEAR et_table.
*--------------------------------------------------------------------*
* Handle each Row until end on right side
*--------------------------------------------------------------------*
            WHILE lv_actual_row <= lv_max_row .

*--------------------------------------------------------------------*
* Handle each Column until end on bottom
* First step is to step back on first column
*--------------------------------------------------------------------*
              lv_actual_col =  iv_skipped_cols + 1.

              UNASSIGN <ls_line>.
              APPEND INITIAL LINE TO et_table ASSIGNING <ls_line>.
              IF sy-subrc <> 0 OR <ls_line> IS NOT ASSIGNED.
                lv_errormessage = 'Error at inserting new Line to internal Table'(002).
                zcx_excel=>raise_text( lv_errormessage ).
              ENDIF.
              WHILE lv_actual_col <= lv_max_col.

                lv_delta_col = lv_actual_col - iv_skipped_cols.
                ASSIGN COMPONENT lv_delta_col OF STRUCTURE <ls_line> TO <lv_value>.
                IF sy-subrc <> 0.
                  lv_actual_col_string = lv_actual_col.
                  lv_actual_row_string = lv_actual_row.
                  CONCATENATE 'Error at assigning field (Col:'(004) lv_actual_col_string ' Row:'(005) lv_actual_row_string INTO lv_errormessage.
                  zcx_excel=>raise_text( lv_errormessage ).
                ENDIF.

                me->get_cell(
                  EXPORTING
                    ip_column  = lv_actual_col    " Cell Column
                    ip_row     = lv_actual_row    " Cell Row
                  IMPORTING
                    ep_value   = lv_value    " Cell Value
                    ep_rc      = lv_rc    " Return Value of ABAP Statements
                ).
                IF lv_rc <> 0
                  AND lv_rc <> 4                                                   "No found error means, zero/no value in cell
                  AND lv_rc <> 8. "rc is 8 when the last row contains cells with zero / no values
                  lv_actual_col_string = lv_actual_col.
                  lv_actual_row_string = lv_actual_row.
                  CONCATENATE 'Error at reading field value (Col:'(007) lv_actual_col_string ' Row:'(005) lv_actual_row_string INTO lv_errormessage.
                  zcx_excel=>raise_text( lv_errormessage ).
                ENDIF.

                TRY.
                    DESCRIBE FIELD <lv_value> TYPE lv_type.
                    IF lv_type = 'D'.
                      <lv_value> = zcl_excel_common=>excel_string_to_date( ip_value = lv_value ).
                    ELSE.
                      <lv_value> = lv_value. "Will raise exception if data type of <lv_value> is not float (or decfloat16/34) and excel delivers exponential number e.g. -2.9398924194538267E-2
                    ENDIF.
                  CATCH cx_sy_conversion_error INTO lx_conversion_error.
                    "Another try with conversion to float...
                    IF lv_type = 'P'.
                      <lv_value> = lv_float = lv_value.
                    ELSE.
                      RAISE EXCEPTION lx_conversion_error. "Pass on original exception
                    ENDIF.
                ENDTRY.

*  CATCH zcx_excel.    "
                ADD 1 TO lv_actual_col.
              ENDWHILE.
              ADD 1 TO lv_actual_row.
            ENDWHILE.

            IF iv_skip_bottom_empty_rows = abap_true.
              lv_tabix = lines( et_table ).
              WHILE lv_tabix >= 1.
                READ TABLE et_table INDEX lv_tabix ASSIGNING <ls_line>.
                ASSERT sy-subrc = 0.
                IF <ls_line> IS NOT INITIAL.
                  EXIT.
                ENDIF.
                DELETE et_table INDEX lv_tabix.
                lv_tabix = lv_tabix - 1.
              ENDWHILE.
            ENDIF.

          ENDIF.


        ENDIF.

      CATCH cx_sy_assign_cast_illegal_cast.
        lv_actual_col_string = lv_actual_col.
        lv_actual_row_string = lv_actual_row.
        CONCATENATE 'Error at assigning field (Col:'(004) lv_actual_col_string ' Row:'(005) lv_actual_row_string INTO lv_errormessage.
        zcx_excel=>raise_text( lv_errormessage ).
      CATCH cx_sy_assign_cast_unknown_type.
        lv_actual_col_string = lv_actual_col.
        lv_actual_row_string = lv_actual_row.
        CONCATENATE 'Error at assigning field (Col:'(004) lv_actual_col_string ' Row:'(005) lv_actual_row_string INTO lv_errormessage.
        zcx_excel=>raise_text( lv_errormessage ).
      CATCH cx_sy_assign_out_of_range.
        lv_errormessage = 'Internal table has less columns than excel'(003).
        zcx_excel=>raise_text( lv_errormessage ).
      CATCH cx_sy_conversion_error.
        lv_actual_col_string = lv_actual_col.
        lv_actual_row_string = lv_actual_row.
        CONCATENATE 'Error at converting field value (Col:'(006) lv_actual_col_string ' Row:'(005) lv_actual_row_string INTO lv_errormessage.
        zcx_excel=>raise_text( lv_errormessage ).

    ENDTRY.
  ENDMETHOD.                    "get_table


  METHOD get_tables_iterator.
    eo_iterator = tables->get_iterator( ).
  ENDMETHOD.                    "GET_TABLES_ITERATOR


  METHOD get_tables_size.
    ep_size = tables->size( ).
  ENDMETHOD.                    "GET_TABLES_SIZE


  METHOD get_title.
    DATA lv_value TYPE string.
    IF ip_escaped EQ abap_true.
      lv_value = me->title.
      ep_title = zcl_excel_common=>escape_string( lv_value ).
    ELSE.
      ep_title = me->title.
    ENDIF.
  ENDMETHOD.                    "GET_TITLE


  METHOD get_value_type.
    DATA: lo_addit    TYPE REF TO cl_abap_elemdescr,
          ls_dfies    TYPE dfies,
          l_function  TYPE funcname,
          l_value(50) TYPE c.

    ep_value = ip_value.
    ep_value_type = cl_abap_typedescr=>typekind_string. " Thats our default if something goes wrong.

    TRY.
        lo_addit            ?= cl_abap_typedescr=>describe_by_data( ip_value ).
      CATCH cx_sy_move_cast_error.
        CLEAR lo_addit.
    ENDTRY.
    IF lo_addit IS BOUND.
      lo_addit->get_ddic_field( RECEIVING  p_flddescr   = ls_dfies
                                EXCEPTIONS not_found    = 1
                                           no_ddic_type = 2
                                           OTHERS       = 3 ) .
      IF sy-subrc = 0.
        ep_value_type = ls_dfies-inttype.

        IF ls_dfies-convexit IS NOT INITIAL.
* We need to convert with output conversion function
          CONCATENATE 'CONVERSION_EXIT_' ls_dfies-convexit '_OUTPUT' INTO l_function.
          SELECT SINGLE funcname INTO l_function
                FROM tfdir
                WHERE funcname = l_function.
          IF sy-subrc = 0.
            CALL FUNCTION l_function
              EXPORTING
                input  = ip_value
              IMPORTING
                output = l_value
              EXCEPTIONS
                OTHERS = 1.
            IF sy-subrc <> 0.
            ELSE.
              TRY.
                  ep_value = l_value.
                CATCH cx_root.
                  ep_value = ip_value.
              ENDTRY.
            ENDIF.
          ENDIF.
        ENDIF.
      ELSE.
        ep_value_type = lo_addit->get_data_type_kind( ip_value ).
      ENDIF.
    ENDIF.

  ENDMETHOD.                    "GET_VALUE_TYPE


  METHOD is_cell_merged.

    DATA: lv_column TYPE i.

    FIELD-SYMBOLS: <ls_merged_cell> LIKE LINE OF me->mt_merged_cells.

    lv_column = zcl_excel_common=>convert_column2int( ip_column ).

    rp_is_merged = abap_false.                                        " Assume not in merged area

    LOOP AT me->mt_merged_cells ASSIGNING <ls_merged_cell>.

      IF    <ls_merged_cell>-col_from <= lv_column
        AND <ls_merged_cell>-col_to   >= lv_column
        AND <ls_merged_cell>-row_from <= ip_row
        AND <ls_merged_cell>-row_to   >= ip_row.
        rp_is_merged = abap_true.                                     " until we are proven different
        RETURN.
      ENDIF.

    ENDLOOP.

  ENDMETHOD.                    "IS_CELL_MERGED


  METHOD move_supplied_borders.

    DATA: ls_borderx TYPE zexcel_s_cstylex_border.

    IF iv_border_supplied = abap_true.  " only act if parameter was supplied
      IF iv_xborder_supplied = abap_true. "
        ls_borderx = is_xborder.             " use supplied x-parameter
      ELSE.
        CLEAR ls_borderx WITH 'X'. " <============================== DDIC structure enh. category to set?
        " clear in a way that would be expected to work easily
        IF is_border-border_style IS  INITIAL.
          CLEAR ls_borderx-border_style.
        ENDIF.
        clear_initial_colorxfields(
          EXPORTING
            is_color  = is_border-border_color
          CHANGING
            cs_xcolor = ls_borderx-border_color ).
      ENDIF.
      MOVE-CORRESPONDING is_border  TO cs_complete_style_border.
      MOVE-CORRESPONDING ls_borderx TO cs_complete_stylex_border.
    ENDIF.

  ENDMETHOD.


  METHOD normalize_column_heading_texts.

    DATA: lt_field_catalog      TYPE zexcel_t_fieldcatalog,
          lv_value_lowercase    TYPE string,
          lv_scrtext_l_initial  TYPE zexcel_column_name,
          lv_long_text          TYPE string,
          lv_max_length         TYPE i,
          lv_temp_length        TYPE i,
          lv_syindex            TYPE c LENGTH 3,
          lt_column_name_buffer TYPE SORTED TABLE OF string WITH UNIQUE KEY table_line.
    FIELD-SYMBOLS: <ls_field_catalog> TYPE zexcel_s_fieldcatalog,
                   <scrtxt1>          TYPE any,
                   <scrtxt2>          TYPE any,
                   <scrtxt3>          TYPE any.

    " Due to restrictions in new table object we cannot have two columns with the same name
    " Check if a column with the same name exists, if exists add a counter
    " If no medium description is provided we try to use small or long

    lt_field_catalog = it_field_catalog.

    LOOP AT lt_field_catalog ASSIGNING <ls_field_catalog> WHERE dynpfld EQ abap_true.

      IF <ls_field_catalog>-column_name IS INITIAL.

        CASE iv_default_descr.
          WHEN 'M'.
            ASSIGN <ls_field_catalog>-scrtext_m TO <scrtxt1>.
            ASSIGN <ls_field_catalog>-scrtext_s TO <scrtxt2>.
            ASSIGN <ls_field_catalog>-scrtext_l TO <scrtxt3>.
          WHEN 'S'.
            ASSIGN <ls_field_catalog>-scrtext_s TO <scrtxt1>.
            ASSIGN <ls_field_catalog>-scrtext_m TO <scrtxt2>.
            ASSIGN <ls_field_catalog>-scrtext_l TO <scrtxt3>.
          WHEN 'L'.
            ASSIGN <ls_field_catalog>-scrtext_l TO <scrtxt1>.
            ASSIGN <ls_field_catalog>-scrtext_m TO <scrtxt2>.
            ASSIGN <ls_field_catalog>-scrtext_s TO <scrtxt3>.
          WHEN OTHERS.
            ASSIGN <ls_field_catalog>-scrtext_m TO <scrtxt1>.
            ASSIGN <ls_field_catalog>-scrtext_s TO <scrtxt2>.
            ASSIGN <ls_field_catalog>-scrtext_l TO <scrtxt3>.
        ENDCASE.

        IF <scrtxt1> IS NOT INITIAL.
          <ls_field_catalog>-column_name = <scrtxt1>.
        ELSEIF <scrtxt2> IS NOT INITIAL.
          <ls_field_catalog>-column_name = <scrtxt2>.
        ELSEIF <scrtxt3> IS NOT INITIAL.
          <ls_field_catalog>-column_name = <scrtxt3>.
        ELSE.
          <ls_field_catalog>-column_name = 'Column'.  " default value as Excel does
        ENDIF.
      ENDIF.

      lv_scrtext_l_initial = <ls_field_catalog>-column_name.
      DESCRIBE FIELD <ls_field_catalog>-column_name LENGTH lv_max_length IN CHARACTER MODE.
      DO.
        lv_value_lowercase = <ls_field_catalog>-column_name.
        TRANSLATE lv_value_lowercase TO LOWER CASE.
        READ TABLE lt_column_name_buffer TRANSPORTING NO FIELDS WITH KEY table_line = lv_value_lowercase BINARY SEARCH.
        IF sy-subrc <> 0.
          INSERT lv_value_lowercase INTO TABLE lt_column_name_buffer.
          EXIT.
        ELSE.
          lv_syindex = sy-index.
          CONCATENATE lv_scrtext_l_initial lv_syindex INTO lv_long_text.
          IF strlen( lv_long_text ) <= lv_max_length.
            <ls_field_catalog>-column_name = lv_long_text.
          ELSE.
            lv_temp_length = strlen( lv_scrtext_l_initial ) - 1.
            lv_scrtext_l_initial = substring( val = lv_scrtext_l_initial len = lv_temp_length ).
            CONCATENATE lv_scrtext_l_initial lv_syindex INTO <ls_field_catalog>-column_name.
          ENDIF.
        ENDIF.
      ENDDO.

    ENDLOOP.

    result = lt_field_catalog.

  ENDMETHOD.


  METHOD normalize_columnrow_parameter.

    IF ( ( ip_column IS NOT INITIAL OR ip_row IS NOT INITIAL ) AND ip_columnrow IS NOT INITIAL )
        OR ( ip_column IS INITIAL AND ip_row IS INITIAL AND ip_columnrow IS INITIAL ).
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'Please provide either row and column, or cell reference'.
    ENDIF.

    IF ip_columnrow IS NOT INITIAL.
      zcl_excel_common=>convert_columnrow2column_a_row(
        EXPORTING
          i_columnrow  = ip_columnrow
        IMPORTING
          e_column_int = ep_column
          e_row        = ep_row ).
    ELSE.
      ep_column = zcl_excel_common=>convert_column2int( ip_column ).
      ep_row    = ip_row.
    ENDIF.

  ENDMETHOD.


  METHOD normalize_range_parameter.

    DATA: lv_errormessage TYPE string.

    IF ( ( ip_column_start IS NOT INITIAL OR ip_column_end IS NOT INITIAL
            OR ip_row IS NOT INITIAL OR ip_row_to IS NOT INITIAL ) AND ip_range IS NOT INITIAL )
        OR ( ip_column_start IS INITIAL AND ip_column_end IS INITIAL
            AND ip_row IS INITIAL AND ip_row_to IS INITIAL AND ip_range IS INITIAL ).
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'Please provide either row and column interval, or range reference'.
    ENDIF.

    IF ip_range IS NOT INITIAL.
      zcl_excel_common=>convert_range2column_a_row(
        EXPORTING
          i_range            = ip_range
        IMPORTING
          e_column_start_int = ep_column_start
          e_column_end_int   = ep_column_end
          e_row_start        = ep_row
          e_row_end          = ep_row_to ).
    ELSE.
      IF ip_column_start IS INITIAL.
        ep_column_start = zcl_excel_common=>c_excel_sheet_min_col.
      ELSE.
        ep_column_start = zcl_excel_common=>convert_column2int( ip_column_start ).
      ENDIF.
      IF ip_column_end IS INITIAL.
        ep_column_end = ep_column_start.
      ELSE.
        ep_column_end = zcl_excel_common=>convert_column2int( ip_column_end ).
      ENDIF.
      ep_row = ip_row.
      IF ep_row IS INITIAL.
        ep_row = zcl_excel_common=>c_excel_sheet_min_row.
      ENDIF.
      ep_row_to = ip_row_to.
      IF ep_row_to IS INITIAL.
        ep_row_to = ep_row.
      ENDIF.
    ENDIF.

    IF ep_row > ep_row_to.
      lv_errormessage = 'First row larger than last row'(405).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

    IF ep_column_start > ep_column_end.
      lv_errormessage = 'First column larger than last column'(406).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

  ENDMETHOD.


  METHOD normalize_style_parameter.

    DATA: lo_style_type TYPE REF TO cl_abap_typedescr.
    FIELD-SYMBOLS:
      <style> TYPE REF TO zcl_excel_style.

    CHECK ip_style_or_guid IS NOT INITIAL.

    lo_style_type = cl_abap_typedescr=>describe_by_data( ip_style_or_guid ).
    IF lo_style_type->type_kind = lo_style_type->typekind_oref.
      lo_style_type = cl_abap_typedescr=>describe_by_object_ref( ip_style_or_guid ).
      IF lo_style_type->absolute_name = '\CLASS=ZCL_EXCEL_STYLE'.
        ASSIGN ip_style_or_guid TO <style>.
        rv_guid = <style>->get_guid( ).
      ENDIF.

    ELSEIF lo_style_type->absolute_name = '\TYPE=ZEXCEL_CELL_STYLE'.
      rv_guid = ip_style_or_guid.

    ELSEIF lo_style_type->type_kind = lo_style_type->typekind_hex.
      rv_guid = ip_style_or_guid.

    ELSE.
      RAISE EXCEPTION TYPE zcx_excel EXPORTING error = 'IP_GUID type must be either REF TO zcl_excel_style or zexcel_cell_style'.
    ENDIF.

  ENDMETHOD.


  METHOD print_title_set_range.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmoecker,                            2012-12-02
*--------------------------------------------------------------------*


    DATA: lo_range_iterator         TYPE REF TO zcl_excel_collection_iterator,
          lo_range                  TYPE REF TO zcl_excel_range,
          lv_repeat_range_sheetname TYPE string,
          lv_repeat_range_col       TYPE string,
          lv_row_char_from          TYPE char10,
          lv_row_char_to            TYPE char10,
          lv_repeat_range_row       TYPE string,
          lv_repeat_range           TYPE string.


*--------------------------------------------------------------------*
* Get range that represents printarea
* if non-existant, create it
*--------------------------------------------------------------------*
    lo_range_iterator = me->get_ranges_iterator( ).
    WHILE lo_range_iterator->has_next( ) = abap_true.

      lo_range ?= lo_range_iterator->get_next( ).
      IF lo_range->name = zif_excel_sheet_printsettings=>gcv_print_title_name.
        EXIT.  " Found it
      ENDIF.
      CLEAR lo_range.

    ENDWHILE.


    IF me->print_title_col_from IS INITIAL AND
       me->print_title_row_from IS INITIAL.
*--------------------------------------------------------------------*
* No print titles are present,
*--------------------------------------------------------------------*
      IF lo_range IS BOUND.
        me->ranges->remove( lo_range ).
      ENDIF.
    ELSE.
*--------------------------------------------------------------------*
* Print titles are present,
*--------------------------------------------------------------------*
      IF lo_range IS NOT BOUND.
        lo_range =  me->add_new_range( ).
        lo_range->name = zif_excel_sheet_printsettings=>gcv_print_title_name.
      ENDIF.

      lv_repeat_range_sheetname = me->get_title( ).
      lv_repeat_range_sheetname = zcl_excel_common=>escape_string( lv_repeat_range_sheetname ).

*--------------------------------------------------------------------*
* Repeat-columns
*--------------------------------------------------------------------*
      IF me->print_title_col_from IS NOT INITIAL.
        CONCATENATE lv_repeat_range_sheetname
                    '!$' me->print_title_col_from
                    ':$' me->print_title_col_to
            INTO lv_repeat_range_col.
      ENDIF.

*--------------------------------------------------------------------*
* Repeat-rows
*--------------------------------------------------------------------*
      IF me->print_title_row_from IS NOT INITIAL.
        lv_row_char_from = me->print_title_row_from.
        lv_row_char_to   = me->print_title_row_to.
        CONCATENATE '!$' lv_row_char_from
                    ':$' lv_row_char_to
            INTO lv_repeat_range_row.
        CONDENSE lv_repeat_range_row NO-GAPS.
        CONCATENATE lv_repeat_range_sheetname
                    lv_repeat_range_row
            INTO lv_repeat_range_row.
      ENDIF.

*--------------------------------------------------------------------*
* Concatenate repeat-rows and columns
*--------------------------------------------------------------------*
      IF lv_repeat_range_col IS INITIAL.
        lv_repeat_range = lv_repeat_range_row.
      ELSEIF lv_repeat_range_row IS INITIAL.
        lv_repeat_range = lv_repeat_range_col.
      ELSE.
        CONCATENATE lv_repeat_range_col lv_repeat_range_row
            INTO lv_repeat_range SEPARATED BY ','.
      ENDIF.


      lo_range->set_range_value( lv_repeat_range ).
    ENDIF.



  ENDMETHOD.                    "PRINT_TITLE_SET_RANGE


  METHOD set_area.

    DATA: lv_row              TYPE zexcel_cell_row,
          lv_row_start        TYPE zexcel_cell_row,
          lv_row_end          TYPE zexcel_cell_row,
          lv_column_int       TYPE zexcel_cell_column,
          lv_column           TYPE zexcel_cell_column_alpha,
          lv_column_start_int TYPE zexcel_cell_column,
          lv_column_end_int   TYPE zexcel_cell_column.

    normalize_range_parameter( EXPORTING ip_range        = ip_range
                                         ip_column_start = ip_column_start     ip_column_end = ip_column_end
                                         ip_row          = ip_row              ip_row_to     = ip_row_to
                               IMPORTING ep_column_start = lv_column_start_int ep_column_end = lv_column_end_int
                                         ep_row          = lv_row_start        ep_row_to     = lv_row_end ).

    " IP_AREA has been added to maintain ascending compatibility (see discussion in PR 869)
    IF ip_merge = abap_true OR ip_area = c_area-topleft.

      IF ip_data_type IS SUPPLIED OR
         ip_abap_type IS SUPPLIED.

        me->set_cell( ip_column    = lv_column_start_int
                      ip_row       = lv_row_start
                      ip_value     = ip_value
                      ip_formula   = ip_formula
                      ip_style     = ip_style
                      ip_hyperlink = ip_hyperlink
                      ip_data_type = ip_data_type
                      ip_abap_type = ip_abap_type ).

      ELSE.

        me->set_cell( ip_column    = lv_column_start_int
                      ip_row       = lv_row_start
                      ip_value     = ip_value
                      ip_formula   = ip_formula
                      ip_style     = ip_style
                      ip_hyperlink = ip_hyperlink ).

      ENDIF.

    ELSE.

      lv_column_int = lv_column_start_int.
      WHILE lv_column_int <= lv_column_end_int.

        lv_column = zcl_excel_common=>convert_column2alpha( lv_column_int ).
        lv_row = lv_row_start.

        WHILE lv_row <= lv_row_end.

          IF ip_data_type IS SUPPLIED OR
             ip_abap_type IS SUPPLIED.

            me->set_cell( ip_column    = lv_column
                          ip_row       = lv_row
                          ip_value     = ip_value
                          ip_formula   = ip_formula
                          ip_style     = ip_style
                          ip_hyperlink = ip_hyperlink
                          ip_data_type = ip_data_type
                          ip_abap_type = ip_abap_type ).

          ELSE.

            me->set_cell( ip_column    = lv_column
                          ip_row       = lv_row
                          ip_value     = ip_value
                          ip_formula   = ip_formula
                          ip_style     = ip_style
                          ip_hyperlink = ip_hyperlink ).

          ENDIF.

          ADD 1 TO lv_row.
        ENDWHILE.

        ADD 1 TO lv_column_int.
      ENDWHILE.

    ENDIF.

    IF ip_style IS SUPPLIED.

      me->set_area_style( ip_column_start = lv_column_start_int
                          ip_column_end   = lv_column_end_int
                          ip_row          = lv_row_start
                          ip_row_to       = lv_row_end
                          ip_style        = ip_style ).
    ENDIF.

    IF ip_merge IS SUPPLIED AND ip_merge = abap_true.

      me->set_merge( ip_column_start = lv_column_start_int
                     ip_column_end   = lv_column_end_int
                     ip_row          = lv_row_start
                     ip_row_to       = lv_row_end ).

    ENDIF.

  ENDMETHOD.                    "set_area


  METHOD set_area_formula.
    DATA: ld_row              TYPE zexcel_cell_row,
          ld_row_start        TYPE zexcel_cell_row,
          ld_row_end          TYPE zexcel_cell_row,
          ld_column           TYPE zexcel_cell_column_alpha,
          ld_column_int       TYPE zexcel_cell_column,
          ld_column_start_int TYPE zexcel_cell_column,
          ld_column_end_int   TYPE zexcel_cell_column.

    normalize_range_parameter( EXPORTING ip_range        = ip_range
                                         ip_column_start = ip_column_start      ip_column_end = ip_column_end
                                         ip_row          = ip_row               ip_row_to     = ip_row_to
                               IMPORTING ep_column_start = ld_column_start_int  ep_column_end = ld_column_end_int
                                         ep_row          = ld_row_start         ep_row_to     = ld_row_end ).

    " IP_AREA has been added to maintain ascending compatibility (see discussion in PR 869)
    IF ip_merge = abap_true OR ip_area = c_area-topleft.

      me->set_cell_formula( ip_column = ld_column_start_int ip_row = ld_row_start
                            ip_formula = ip_formula ).

    ELSE.

      ld_column_int = ld_column_start_int.
      WHILE ld_column_int <= ld_column_end_int.

        ld_column = zcl_excel_common=>convert_column2alpha( ld_column_int ).
        ld_row = ld_row_start.
        WHILE ld_row <= ld_row_end.

          me->set_cell_formula( ip_column = ld_column ip_row = ld_row
                                ip_formula = ip_formula ).

          ADD 1 TO ld_row.
        ENDWHILE.

        ADD 1 TO ld_column_int.
      ENDWHILE.

    ENDIF.

    IF ip_merge IS SUPPLIED AND ip_merge = abap_true.
      me->set_merge( ip_column_start = ld_column_start_int ip_row = ld_row_start
                     ip_column_end   = ld_column_end_int   ip_row_to = ld_row_end ).
    ENDIF.
  ENDMETHOD.                    "set_area_formula


  METHOD set_area_hyperlink.
    DATA: ld_row_start        TYPE zexcel_cell_row,
          ld_row_end          TYPE zexcel_cell_row,
          ld_column_int       TYPE zexcel_cell_column,
          ld_column_start_int TYPE zexcel_cell_column,
          ld_column_end_int   TYPE zexcel_cell_column,
          ld_current_column   TYPE zexcel_cell_column_alpha,
          ld_current_row      TYPE zexcel_cell_row,
          ld_value            TYPE string,
          ld_formula          TYPE string.
    DATA: lo_hyperlink TYPE REF TO zcl_excel_hyperlink.

    normalize_range_parameter( EXPORTING ip_range        = ip_range
                                         ip_column_start = ip_column_start      ip_column_end = ip_column_end
                                         ip_row          = ip_row               ip_row_to     = ip_row_to
                               IMPORTING ep_column_start = ld_column_start_int  ep_column_end = ld_column_end_int
                                         ep_row          = ld_row_start         ep_row_to     = ld_row_end ).

    ld_column_int = ld_column_start_int.
    WHILE ld_column_int <= ld_column_end_int.
      ld_current_column = zcl_excel_common=>convert_column2alpha( ld_column_int ).
      ld_current_row = ld_row_start.
      WHILE ld_current_row <= ld_row_end.

        me->get_cell( EXPORTING ip_column  = ld_current_column ip_row = ld_current_row
                      IMPORTING ep_value   = ld_value
                                ep_formula = ld_formula ).

        IF ip_is_internal = abap_true.
          lo_hyperlink = zcl_excel_hyperlink=>create_internal_link( iv_location = ip_url ).
        ELSE.
          lo_hyperlink = zcl_excel_hyperlink=>create_external_link( iv_url = ip_url ).
        ENDIF.

        me->set_cell( ip_column = ld_current_column ip_row = ld_current_row ip_value = ld_value ip_formula = ld_formula ip_hyperlink = lo_hyperlink ).

        ADD 1 TO ld_current_row.
      ENDWHILE.
      ADD 1 TO ld_column_int.
    ENDWHILE.

  ENDMETHOD.                    "SET_AREA_HYPERLINK


  METHOD set_area_style.
    DATA: ld_row_start        TYPE zexcel_cell_row,
          ld_row_end          TYPE zexcel_cell_row,
          ld_column_int       TYPE zexcel_cell_column,
          ld_column_start_int TYPE zexcel_cell_column,
          ld_column_end_int   TYPE zexcel_cell_column,
          ld_current_column   TYPE zexcel_cell_column_alpha,
          ld_current_row      TYPE zexcel_cell_row.

    normalize_range_parameter( EXPORTING ip_range        = ip_range
                                         ip_column_start = ip_column_start      ip_column_end = ip_column_end
                                         ip_row          = ip_row               ip_row_to     = ip_row_to
                               IMPORTING ep_column_start = ld_column_start_int  ep_column_end = ld_column_end_int
                                         ep_row          = ld_row_start         ep_row_to     = ld_row_end ).

    ld_column_int = ld_column_start_int.
    WHILE ld_column_int <= ld_column_end_int.
      ld_current_column = zcl_excel_common=>convert_column2alpha( ld_column_int ).
      ld_current_row = ld_row_start.
      WHILE ld_current_row <= ld_row_end.
        me->set_cell_style( ip_row = ld_current_row ip_column = ld_current_column
                            ip_style = ip_style ).
        ADD 1 TO ld_current_row.
      ENDWHILE.
      ADD 1 TO ld_column_int.
    ENDWHILE.
    IF ip_merge IS SUPPLIED AND ip_merge = abap_true.
      me->set_merge( ip_column_start = ld_column_start_int ip_row = ld_row_start
                     ip_column_end   = ld_column_end_int   ip_row_to = ld_row_end ).
    ENDIF.
  ENDMETHOD.                    "SET_AREA_STYLE


  METHOD set_cell.

    DATA: lv_column        TYPE zexcel_cell_column,
          ls_sheet_content TYPE zexcel_s_cell_data,
          lv_row           TYPE zexcel_cell_row,
          lv_value         TYPE zexcel_cell_value,
          lv_data_type     TYPE zexcel_cell_data_type,
          lv_value_type    TYPE abap_typekind,
          lv_style_guid    TYPE zexcel_cell_style,
          lo_addit         TYPE REF TO cl_abap_elemdescr,
          lt_rtf           TYPE zexcel_t_rtf,
          lo_value         TYPE REF TO data,
          lo_value_new     TYPE REF TO data.

    FIELD-SYMBOLS: <fs_sheet_content> TYPE zexcel_s_cell_data,
                   <fs_numeric>       TYPE numeric,
                   <fs_date>          TYPE d,
                   <fs_time>          TYPE t,
                   <fs_value>         TYPE simple,
                   <fs_typekind_int8> TYPE abap_typekind.
    FIELD-SYMBOLS: <fs_column_formula> TYPE mty_s_column_formula.
    FIELD-SYMBOLS: <ls_fieldcat>       TYPE zexcel_s_fieldcatalog.

    IF ip_value  IS NOT SUPPLIED
        AND ip_formula IS NOT SUPPLIED
        AND ip_column_formula_id = 0.
      zcx_excel=>raise_text( 'Please provide the value or formula' ).
    ENDIF.

    normalize_columnrow_parameter( EXPORTING ip_columnrow = ip_columnrow
                                             ip_column    = ip_column
                                             ip_row       = ip_row
                                   IMPORTING ep_column    = lv_column
                                             ep_row       = lv_row ).

* Begin of change issue #152 - don't touch exisiting style if only value is passed
    IF ip_column_formula_id <> 0.
      check_cell_column_formula(
          it_column_formulas   = column_formulas
          ip_column_formula_id = ip_column_formula_id
          ip_formula           = ip_formula
          ip_value             = ip_value
          ip_row               = lv_row
          ip_column            = lv_column ).
    ENDIF.
    READ TABLE sheet_content ASSIGNING <fs_sheet_content> WITH TABLE KEY cell_row    = lv_row      " Changed to access via table key , Stefan Schmöcker, 2013-08-03
                                                                         cell_column = lv_column.
    IF sy-subrc = 0.
      IF ip_style IS INITIAL.
        " If no style is provided as method-parameter and cell is found use cell's current style
        lv_style_guid = <fs_sheet_content>-cell_style.
      ELSE.
        " Style provided as method-parameter --> use this
        lv_style_guid = normalize_style_parameter( ip_style ).
      ENDIF.
    ELSE.
      " No cell found --> use supplied style even if empty
      lv_style_guid = normalize_style_parameter( ip_style ).
    ENDIF.
* End of change issue #152 - don't touch exisiting style if only value is passed

    IF ip_value IS SUPPLIED.
      "if data type is passed just write the value. Otherwise map abap type to excel and perform conversion
      "IP_DATA_TYPE is passed by excel reader so source types are preserved
*First we get reference into local var.
      IF ip_conv_exit_length = abap_true.
        lo_value = create_data_conv_exit_length( ip_value ).
      ELSE.
        CREATE DATA lo_value LIKE ip_value.
      ENDIF.
      ASSIGN lo_value->* TO <fs_value>.
      <fs_value> = ip_value.
      IF ip_data_type IS SUPPLIED.
        IF ip_abap_type IS NOT SUPPLIED.
          get_value_type( EXPORTING ip_value      = ip_value
                          IMPORTING ep_value      = <fs_value> ) .
        ENDIF.
        lv_value = <fs_value>.
        lv_data_type = ip_data_type.
      ELSE.
        IF ip_abap_type IS SUPPLIED.
          lv_value_type = ip_abap_type.
        ELSE.
          get_value_type( EXPORTING ip_value      = ip_value
                          IMPORTING ep_value      = <fs_value>
                                    ep_value_type = lv_value_type ).
        ENDIF.

        ASSIGN ('CL_ABAP_TYPEDESCR=>TYPEKIND_INT8') TO <fs_typekind_int8>.
        IF sy-subrc <> 0.
          ASSIGN space TO <fs_typekind_int8>. "not used as typekind!
        ENDIF.

        CASE lv_value_type.
          WHEN cl_abap_typedescr=>typekind_int OR cl_abap_typedescr=>typekind_int1 OR cl_abap_typedescr=>typekind_int2
            OR <fs_typekind_int8>. "Allow INT8 types columns
            IF lv_value_type = <fs_typekind_int8>.
              CALL METHOD cl_abap_elemdescr=>('GET_INT8') RECEIVING p_result = lo_addit.
            ELSE.
              lo_addit = cl_abap_elemdescr=>get_i( ).
            ENDIF.
            CREATE DATA lo_value_new TYPE HANDLE lo_addit.
            ASSIGN lo_value_new->* TO <fs_numeric>.
            IF sy-subrc = 0.
              <fs_numeric> = <fs_value>.
              lv_value = zcl_excel_common=>number_to_excel_string( ip_value = <fs_numeric> ).
            ENDIF.

          WHEN cl_abap_typedescr=>typekind_float OR cl_abap_typedescr=>typekind_packed OR
               cl_abap_typedescr=>typekind_decfloat OR
               cl_abap_typedescr=>typekind_decfloat16 OR
               cl_abap_typedescr=>typekind_decfloat34.
            IF lv_value_type = cl_abap_typedescr=>typekind_packed
                AND ip_currency IS NOT INITIAL.
              lv_value = zcl_excel_common=>number_to_excel_string( ip_value    = <fs_value>
                                                                   ip_currency = ip_currency ).
            ELSE.
              lo_addit = cl_abap_elemdescr=>get_f( ).
              CREATE DATA lo_value_new TYPE HANDLE lo_addit.
              ASSIGN lo_value_new->* TO <fs_numeric>.
              IF sy-subrc = 0.
                <fs_numeric> = <fs_value>.
                lv_value = zcl_excel_common=>number_to_excel_string( ip_value = <fs_numeric> ).
              ENDIF.
            ENDIF.

          WHEN cl_abap_typedescr=>typekind_char OR cl_abap_typedescr=>typekind_string OR cl_abap_typedescr=>typekind_num OR
               cl_abap_typedescr=>typekind_hex.
            lv_value = <fs_value>.
            lv_data_type = 's'.

          WHEN cl_abap_typedescr=>typekind_date.
            lo_addit = cl_abap_elemdescr=>get_d( ).
            CREATE DATA lo_value_new TYPE HANDLE lo_addit.
            ASSIGN lo_value_new->* TO <fs_date>.
            IF sy-subrc = 0.
              <fs_date> = <fs_value>.
              lv_value = zcl_excel_common=>date_to_excel_string( ip_value = <fs_date> ) .
            ENDIF.
* Begin of change issue #152 - don't touch exisiting style if only value is passed
* Moved to end of routine - apply date-format even if other styleinformation is passed
*          IF ip_style IS NOT SUPPLIED. "get default date format in case parameter is initial
*            lo_style = excel->add_new_style( ).
*            lo_style->number_format->format_code = get_default_excel_date_format( ).
*            lv_style_guid = lo_style->get_guid( ).
*          ENDIF.
* End of change issue #152 - don't touch exisiting style if only value is passed

          WHEN cl_abap_typedescr=>typekind_time.
            lo_addit = cl_abap_elemdescr=>get_t( ).
            CREATE DATA lo_value_new TYPE HANDLE lo_addit.
            ASSIGN lo_value_new->* TO <fs_time>.
            IF sy-subrc = 0.
              <fs_time> = <fs_value>.
              lv_value = zcl_excel_common=>time_to_excel_string( ip_value = <fs_time> ).
            ENDIF.
* Begin of change issue #152 - don't touch exisiting style if only value is passed
* Moved to end of routine - apply time-format even if other styleinformation is passed
*          IF ip_style IS NOT SUPPLIED. "get default time format for user in case parameter is initial
*            lo_style = excel->add_new_style( ).
*            lo_style->number_format->format_code = zcl_excel_style_number_format=>c_format_date_time6.
*            lv_style_guid = lo_style->get_guid( ).
*          ENDIF.
* End of change issue #152 - don't touch exisiting style if only value is passed

          WHEN OTHERS.
            zcx_excel=>raise_text( 'Invalid data type of input value' ).
        ENDCASE.
      ENDIF.

      IF <fs_sheet_content> IS ASSIGNED AND <fs_sheet_content>-table_header IS NOT INITIAL AND lv_value IS NOT INITIAL.
        READ TABLE <fs_sheet_content>-table->fieldcat ASSIGNING <ls_fieldcat> WITH KEY fieldname = <fs_sheet_content>-table_fieldname.
        IF sy-subrc = 0.
          <ls_fieldcat>-column_name = lv_value.
          IF <ls_fieldcat>-column_name <> lv_value.
            zcx_excel=>raise_text( 'Cell is table column header - this value is not allowed' ).
          ENDIF.
        ENDIF.
      ENDIF.

    ENDIF.

    IF ip_hyperlink IS BOUND.
      ip_hyperlink->set_cell_reference( ip_column = lv_column
                                        ip_row = lv_row ).
      me->hyperlinks->add( ip_hyperlink ).
    ENDIF.

    IF lv_value CS '_x'.
      " Issue #761 value "_x0041_" rendered as "A".
      " "_x...._", where "." is 0-9 a-f or A-F (case insensitive), is an internal value in sharedStrings.xml
      " that Excel uses to store special characters, it's interpreted like Unicode character U+....
      " for instance "_x0041_" is U+0041 which is "A".
      " To not interpret such text, the first underscore is replaced with "_x005f_".
      " The value "_x0041_" is to be stored internally "_x005f_x0041_" so that it's rendered like "_x0041_".
      " Note that REGEX is time consuming, it's why "CS" is used above to improve the performance.
      REPLACE ALL OCCURRENCES OF REGEX '_(x[0-9a-fA-F]{4}_)' IN lv_value WITH '_x005f_$1' RESPECTING CASE.
    ENDIF.

* Begin of change issue #152 - don't touch exisiting style if only value is passed
* Read table moved up, so that current style may be evaluated

    IF <fs_sheet_content> IS ASSIGNED.
* End of change issue #152 - don't touch exisiting style if only value is passed
      <fs_sheet_content>-cell_value   = lv_value.
      <fs_sheet_content>-cell_formula = ip_formula.
      <fs_sheet_content>-column_formula_id = ip_column_formula_id.
      <fs_sheet_content>-cell_style   = lv_style_guid.
      <fs_sheet_content>-data_type    = lv_data_type.
    ELSE.
      ls_sheet_content-cell_row     = lv_row.
      ls_sheet_content-cell_column  = lv_column.
      ls_sheet_content-cell_value   = lv_value.
      ls_sheet_content-cell_formula = ip_formula.
      ls_sheet_content-column_formula_id = ip_column_formula_id.
      ls_sheet_content-cell_style   = lv_style_guid.
      ls_sheet_content-data_type    = lv_data_type.
      ls_sheet_content-cell_coords  = zcl_excel_common=>convert_column_a_row2columnrow( i_column = lv_column i_row = lv_row ).
      INSERT ls_sheet_content INTO TABLE sheet_content ASSIGNING <fs_sheet_content>. "ins #152 - Now <fs_sheet_content> always holds the data

    ENDIF.

    IF ip_formula IS INITIAL AND lv_value IS NOT INITIAL AND it_rtf IS NOT INITIAL.
      lt_rtf = it_rtf.
      check_rtf( EXPORTING ip_value = lv_value
                           ip_style = lv_style_guid
                 CHANGING  ct_rtf   = lt_rtf ).
      <fs_sheet_content>-rtf_tab = lt_rtf.
    ENDIF.

* Begin of change issue #152 - don't touch exisiting style if only value is passed
* For Date- or Timefields change the formatcode if nothing is set yet
* Enhancement option:  Check if existing formatcode is a date/ or timeformat
*                      If not, use default
    DATA: lo_format_code_datetime TYPE zexcel_number_format.
    DATA: stylemapping    TYPE zexcel_s_stylemapping.
    IF <fs_sheet_content>-cell_style IS INITIAL.
      <fs_sheet_content>-cell_style = me->excel->get_default_style( ).
    ENDIF.
    CASE lv_value_type.
      WHEN cl_abap_typedescr=>typekind_date.
        TRY.
            stylemapping = me->excel->get_style_to_guid( <fs_sheet_content>-cell_style ).
          CATCH zcx_excel .
        ENDTRY.
        IF stylemapping-complete_stylex-number_format-format_code IS INITIAL OR
           stylemapping-complete_style-number_format-format_code IS INITIAL.
          lo_format_code_datetime = zcl_excel_style_number_format=>c_format_date_std.
        ELSE.
          lo_format_code_datetime = stylemapping-complete_style-number_format-format_code.
        ENDIF.
        me->change_cell_style( ip_column                      = lv_column
                               ip_row                         = lv_row
                               ip_number_format_format_code   = lo_format_code_datetime ).

      WHEN cl_abap_typedescr=>typekind_time.
        TRY.
            stylemapping = me->excel->get_style_to_guid( <fs_sheet_content>-cell_style ).
          CATCH zcx_excel .
        ENDTRY.
        IF stylemapping-complete_stylex-number_format-format_code IS INITIAL OR
           stylemapping-complete_style-number_format-format_code IS INITIAL.
          lo_format_code_datetime = zcl_excel_style_number_format=>c_format_date_time6.
        ELSE.
          lo_format_code_datetime = stylemapping-complete_style-number_format-format_code.
        ENDIF.
        me->change_cell_style( ip_column                      = lv_column
                               ip_row                         = lv_row
                               ip_number_format_format_code   = lo_format_code_datetime ).

    ENDCASE.
* End of change issue #152 - don't touch exisiting style if only value is passed

* Fix issue #162
    lv_value = ip_value.
    IF lv_value CS cl_abap_char_utilities=>cr_lf.
      me->change_cell_style( ip_column               = lv_column
                             ip_row                  = lv_row
                             ip_alignment_wraptext   = abap_true ).
    ENDIF.
* End of Fix issue #162

  ENDMETHOD.                    "SET_CELL


  METHOD set_cell_formula.
    DATA:
      lv_column        TYPE zexcel_cell_column,
      lv_row           TYPE zexcel_cell_row,
      ls_sheet_content LIKE LINE OF me->sheet_content.

    FIELD-SYMBOLS:
                <sheet_content>                 LIKE LINE OF me->sheet_content.

*--------------------------------------------------------------------*
* Get cell to set formula into
*--------------------------------------------------------------------*
    normalize_columnrow_parameter( EXPORTING ip_columnrow = ip_columnrow
                                             ip_column    = ip_column
                                             ip_row       = ip_row
                                   IMPORTING ep_column    = lv_column
                                             ep_row       = lv_row ).

    READ TABLE me->sheet_content ASSIGNING <sheet_content> WITH TABLE KEY cell_row    = lv_row
                                                                          cell_column = lv_column.
    IF sy-subrc <> 0.                   " Create new entry in sheet_content if necessary
      CHECK ip_formula IS NOT INITIAL.  " only create new entry in sheet_content when a formula is passed
      ls_sheet_content-cell_row    = lv_row.
      ls_sheet_content-cell_column = lv_column.
      ls_sheet_content-cell_coords = zcl_excel_common=>convert_column_a_row2columnrow( i_column = lv_column i_row = lv_row ).
      INSERT ls_sheet_content INTO TABLE me->sheet_content ASSIGNING <sheet_content>.
    ENDIF.

*--------------------------------------------------------------------*
* Fieldsymbol now holds the relevant cell
*--------------------------------------------------------------------*
    <sheet_content>-cell_formula = ip_formula.


  ENDMETHOD.                    "SET_CELL_FORMULA


  METHOD set_cell_style.

    DATA: lv_column     TYPE zexcel_cell_column,
          lv_row        TYPE zexcel_cell_row,
          lv_style_guid TYPE zexcel_cell_style.

    FIELD-SYMBOLS: <fs_sheet_content> TYPE zexcel_s_cell_data.

    lv_style_guid = normalize_style_parameter( ip_style ).

    normalize_columnrow_parameter( EXPORTING ip_columnrow = ip_columnrow
                                             ip_column    = ip_column
                                             ip_row       = ip_row
                                   IMPORTING ep_column    = lv_column
                                             ep_row       = lv_row ).

    READ TABLE sheet_content ASSIGNING <fs_sheet_content> WITH KEY cell_row    = lv_row
                                                                   cell_column = lv_column.

    IF sy-subrc EQ 0.
      <fs_sheet_content>-cell_style   = lv_style_guid.
    ELSE.
      set_cell( ip_column = ip_column ip_row = ip_row ip_value = '' ip_style = ip_style ).
    ENDIF.

  ENDMETHOD.                    "SET_CELL_STYLE


  METHOD set_table_reference.

    FIELD-SYMBOLS: <ls_sheet_content> TYPE zexcel_s_cell_data.

    READ TABLE sheet_content ASSIGNING <ls_sheet_content> WITH KEY cell_row    = ip_row
                                                                   cell_column = ip_column.
    IF sy-subrc = 0.
      <ls_sheet_content>-table           = ir_table.
      <ls_sheet_content>-table_fieldname = ip_fieldname.
      <ls_sheet_content>-table_header    = ip_header.
    ELSE.
      zcx_excel=>raise_text( 'Cell not found' ).
    ENDIF.

  ENDMETHOD.


  METHOD set_column_width.
    DATA: lo_column  TYPE REF TO zcl_excel_column.
    DATA: width             TYPE f.

    lo_column = me->get_column( ip_column ).

* if a fix size is supplied use this
    IF ip_width_fix IS SUPPLIED.
      TRY.
          width = ip_width_fix.
          IF width <= 0.
            zcx_excel=>raise_text( 'Please supply a positive number as column-width' ).
          ENDIF.
          lo_column->set_width( width ).
          RETURN.
        CATCH cx_sy_conversion_no_number.
* Strange stuff passed --> raise error
          zcx_excel=>raise_text( 'Unable to interpret supplied input as number' ).
      ENDTRY.
    ENDIF.

* If we get down to here, we have to use whatever is found in autosize.
    lo_column->set_auto_size( ip_width_autosize ).


  ENDMETHOD.                    "SET_COLUMN_WIDTH


  METHOD set_default_excel_date_format.

    IF ip_default_excel_date_format IS INITIAL.
      zcx_excel=>raise_text( 'Default date format cannot be blank' ).
    ENDIF.

    default_excel_date_format = ip_default_excel_date_format.
  ENDMETHOD.                    "SET_DEFAULT_EXCEL_DATE_FORMAT


  METHOD set_ignored_errors.
    mt_ignored_errors = it_ignored_errors.
  ENDMETHOD.


  METHOD set_merge.

    DATA: ls_merge        TYPE mty_merge,
          lv_column_start TYPE zexcel_cell_column,
          lv_column_end   TYPE zexcel_cell_column,
          lv_row          TYPE zexcel_cell_row,
          lv_row_to       TYPE zexcel_cell_row,
          lv_errormessage TYPE string.

    normalize_range_parameter( EXPORTING ip_range        = ip_range
                                         ip_column_start = ip_column_start ip_column_end = ip_column_end
                                         ip_row          = ip_row          ip_row_to     = ip_row_to
                               IMPORTING ep_column_start = lv_column_start ep_column_end = lv_column_end
                                         ep_row          = lv_row          ep_row_to     = lv_row_to ).

    IF ip_value IS SUPPLIED OR ip_formula IS SUPPLIED.
      " if there is a value or formula set the value to the top-left cell
      "maybe it is necessary to support other paramters for set_cell
      IF ip_value IS SUPPLIED.
        me->set_cell( ip_row = lv_row ip_column = lv_column_start
                      ip_value = ip_value ).
      ENDIF.
      IF ip_formula IS SUPPLIED.
        me->set_cell( ip_row = lv_row ip_column = lv_column_start
                      ip_value = ip_formula ).
      ENDIF.
    ENDIF.
    "call to set_merge_style to apply the style to all cells at the matrix
    IF ip_style IS SUPPLIED.
      me->set_merge_style( ip_row = lv_row ip_column_start = lv_column_start
                           ip_row_to = lv_row_to ip_column_end = lv_column_end
                           ip_style = ip_style ).
    ENDIF.
    ...
*--------------------------------------------------------------------*
* Build new range area to insert into range table
*--------------------------------------------------------------------*
    ls_merge-row_from = lv_row.
    ls_merge-row_to   = lv_row_to.
    ls_merge-col_from = lv_column_start.
    ls_merge-col_to   = lv_column_end.

*--------------------------------------------------------------------*
* Check merge not overlapping with existing merges
*--------------------------------------------------------------------*
    LOOP AT me->mt_merged_cells TRANSPORTING NO FIELDS WHERE NOT (    row_from > ls_merge-row_to
                                                                   OR row_to   < ls_merge-row_from
                                                                   OR col_from > ls_merge-col_to
                                                                   OR col_to   < ls_merge-col_from ).
      lv_errormessage = 'Overlapping merges'(404).
      zcx_excel=>raise_text( lv_errormessage ).

    ENDLOOP.

*--------------------------------------------------------------------*
* Everything seems ok --> add to merge table
*--------------------------------------------------------------------*
    INSERT ls_merge INTO TABLE me->mt_merged_cells.

  ENDMETHOD.                    "SET_MERGE


  METHOD set_merge_style.
    DATA: ld_row_start      TYPE zexcel_cell_row,
          ld_row_end        TYPE zexcel_cell_row,
          ld_column_int     TYPE zexcel_cell_column,
          ld_column_start   TYPE zexcel_cell_column,
          ld_column_end     TYPE zexcel_cell_column,
          ld_current_column TYPE zexcel_cell_column_alpha,
          ld_current_row    TYPE zexcel_cell_row.

    normalize_range_parameter( EXPORTING ip_range        = ip_range
                                         ip_column_start = ip_column_start ip_column_end = ip_column_end
                                         ip_row          = ip_row          ip_row_to     = ip_row_to
                               IMPORTING ep_column_start = ld_column_start ep_column_end = ld_column_end
                                         ep_row          = ld_row_start    ep_row_to     = ld_row_end ).

    "set the style cell by cell
    ld_column_int = ld_column_start.
    WHILE ld_column_int <= ld_column_end.
      ld_current_column = zcl_excel_common=>convert_column2alpha( ld_column_int ).
      ld_current_row = ld_row_start.
      WHILE ld_current_row <= ld_row_end.
        me->set_cell_style( ip_row = ld_current_row ip_column = ld_current_column
                            ip_style = ip_style ).
        ADD 1 TO ld_current_row.
      ENDWHILE.
      ADD 1 TO ld_column_int.
    ENDWHILE.
  ENDMETHOD.                    "set_merge_style


  METHOD set_pane_top_left_cell.
    DATA lv_column_int TYPE zexcel_cell_column.
    DATA lv_row TYPE zexcel_cell_row.

    " Validate input value
    zcl_excel_common=>convert_columnrow2column_a_row(
      EXPORTING
        i_columnrow  = iv_columnrow
      IMPORTING
        e_column_int = lv_column_int
        e_row        = lv_row ).
    IF lv_column_int NOT BETWEEN zcl_excel_common=>c_excel_sheet_min_col AND zcl_excel_common=>c_excel_sheet_max_col
        OR lv_row NOT BETWEEN zcl_excel_common=>c_excel_sheet_min_row AND zcl_excel_common=>c_excel_sheet_max_row.
      RAISE EXCEPTION TYPE zcx_excel EXPORTING error = 'Invalid column/row coordinates (valid values: A1 to XFD1048576)'.
    ENDIF.
    pane_top_left_cell = iv_columnrow.
  ENDMETHOD.


  METHOD set_print_gridlines.
    me->print_gridlines = i_print_gridlines.
  ENDMETHOD.                    "SET_PRINT_GRIDLINES


  METHOD set_row_height.
    DATA: lo_row  TYPE REF TO zcl_excel_row.
    DATA: height  TYPE f.

    lo_row = me->get_row( ip_row ).

* if a fix size is supplied use this
    TRY.
        height = ip_height_fix.
        lo_row->set_row_height( height ).
        RETURN.
      CATCH cx_sy_conversion_no_number.
* Strange stuff passed --> raise error
        zcx_excel=>raise_text( 'Unable to interpret supplied input as number' ).
    ENDTRY.

  ENDMETHOD.                    "SET_ROW_HEIGHT


  METHOD set_row_outline.

    DATA: ls_row_outline LIKE LINE OF me->mt_row_outlines.
    FIELD-SYMBOLS: <ls_row_outline> LIKE LINE OF me->mt_row_outlines.

    READ TABLE me->mt_row_outlines ASSIGNING <ls_row_outline> WITH TABLE KEY row_from = iv_row_from
                                                                             row_to   = iv_row_to.
    IF sy-subrc <> 0.
      IF iv_row_from <= 0.
        zcx_excel=>raise_text( 'First row of outline must be a positive number' ).
      ENDIF.
      IF iv_row_to < iv_row_from.
        zcx_excel=>raise_text( 'Last row of outline may not be less than first line of outline' ).
      ENDIF.
      ls_row_outline-row_from = iv_row_from.
      ls_row_outline-row_to   = iv_row_to.
      INSERT ls_row_outline INTO TABLE me->mt_row_outlines ASSIGNING <ls_row_outline>.
    ENDIF.

    CASE iv_collapsed.

      WHEN abap_true
        OR abap_false.
        <ls_row_outline>-collapsed = iv_collapsed.

      WHEN OTHERS.
        zcx_excel=>raise_text( 'Unknown collapse state' ).

    ENDCASE.
  ENDMETHOD.                    "SET_ROW_OUTLINE


  METHOD set_sheetview_top_left_cell.
    DATA lv_column_int TYPE zexcel_cell_column.
    DATA lv_row TYPE zexcel_cell_row.

    " Validate input value
    zcl_excel_common=>convert_columnrow2column_a_row(
      EXPORTING
        i_columnrow  = iv_columnrow
      IMPORTING
        e_column_int = lv_column_int
        e_row        = lv_row ).
    IF lv_column_int NOT BETWEEN zcl_excel_common=>c_excel_sheet_min_col AND zcl_excel_common=>c_excel_sheet_max_col
        OR lv_row NOT BETWEEN zcl_excel_common=>c_excel_sheet_min_row AND zcl_excel_common=>c_excel_sheet_max_row.
      RAISE EXCEPTION TYPE zcx_excel EXPORTING error = 'Invalid column/row coordinates (valid values: A1 to XFD1048576)'.
    ENDIF.
    sheetview_top_left_cell = iv_columnrow.
  ENDMETHOD.


  METHOD set_show_gridlines.
    me->show_gridlines = i_show_gridlines.
  ENDMETHOD.                    "SET_SHOW_GRIDLINES


  METHOD set_show_rowcolheaders.
    me->show_rowcolheaders = i_show_rowcolheaders.
  ENDMETHOD.                    "SET_SHOW_ROWCOLHEADERS


  METHOD set_tabcolor.
    me->tabcolor = iv_tabcolor.
  ENDMETHOD.                    "SET_TABCOLOR


  METHOD set_table.

    DATA: lo_structdescr     TYPE REF TO cl_abap_structdescr,
          lr_data         TYPE REF TO data,
          lt_dfies        TYPE ddfields,
          lv_row_int      TYPE zexcel_cell_row,
          lv_column_int   TYPE zexcel_cell_column,
          lv_column_alpha TYPE zexcel_cell_column_alpha,
          lv_cell_value   TYPE zexcel_cell_value.


    FIELD-SYMBOLS: <fs_table_line> TYPE any,
                   <fs_fldval>     TYPE any,
                   <fs_dfies>      TYPE dfies.

    lv_column_int = zcl_excel_common=>convert_column2int( ip_top_left_column ).
    lv_row_int    = ip_top_left_row.

    CREATE DATA lr_data LIKE LINE OF ip_table.

    lo_structdescr ?= cl_abap_structdescr=>describe_by_data_ref( lr_data ).

    lt_dfies = zcl_excel_common=>describe_structure( io_struct = lo_structdescr ).

* It is better to loop column by column
    LOOP AT lt_dfies ASSIGNING <fs_dfies>.
      lv_column_alpha = zcl_excel_common=>convert_column2alpha( lv_column_int ).

      IF ip_no_header = abap_false.
        " First of all write column header
        lv_cell_value = <fs_dfies>-scrtext_m.
        me->set_cell( ip_column = lv_column_alpha
                      ip_row    = lv_row_int
                      ip_value  = lv_cell_value
                      ip_style  = ip_hdr_style ).
        IF ip_transpose = abap_true.
          ADD 1 TO lv_column_int.
        ELSE.
          ADD 1 TO lv_row_int.
        ENDIF.
      ENDIF.

      LOOP AT ip_table ASSIGNING <fs_table_line>.
        lv_column_alpha = zcl_excel_common=>convert_column2alpha( lv_column_int ).
        ASSIGN COMPONENT <fs_dfies>-fieldname OF STRUCTURE <fs_table_line> TO <fs_fldval>.
        lv_cell_value = <fs_fldval>.
        me->set_cell( ip_column = lv_column_alpha
                      ip_row    = lv_row_int
                      ip_value  = <fs_fldval>   "lv_cell_value
                      ip_style  = ip_body_style ).
        IF ip_transpose = abap_true.
          ADD 1 TO lv_column_int.
        ELSE.
          ADD 1 TO lv_row_int.
        ENDIF.
      ENDLOOP.
      IF ip_transpose = abap_true.
        lv_column_int = zcl_excel_common=>convert_column2int( ip_top_left_column ).
        ADD 1 TO lv_row_int.
      ELSE.
        lv_row_int = ip_top_left_row.
        ADD 1 TO lv_column_int.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.                    "SET_TABLE


  METHOD set_title.
*--------------------------------------------------------------------*
* ToDos:
*        2do §1  The current coding for replacing a named ranges name
*                after renaming a sheet should be checked if it is
*                really working if sheetname should be escaped
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (wip )              2012-12-08
*              - ...
* changes: aligning code
*          message made to support multilinguality
*--------------------------------------------------------------------*
* issue#243 - ' is not allowed as first character in sheet title
*              - Stefan Schmoecker,                          2012-12-02
* changes: added additional check for ' as first character
*--------------------------------------------------------------------*
    DATA: lo_worksheets_iterator TYPE REF TO zcl_excel_collection_iterator,
          lo_worksheet           TYPE REF TO zcl_excel_worksheet,
          errormessage           TYPE string,
          lv_rangesheetname_old  TYPE string,
          lv_rangesheetname_new  TYPE string,
          lo_ranges_iterator     TYPE REF TO zcl_excel_collection_iterator,
          lo_range               TYPE REF TO zcl_excel_range,
          lv_range_value         TYPE zexcel_range_value,
          lv_errormessage        TYPE string.                          " Can't pass '...'(abc) to exception-class


*--------------------------------------------------------------------*
* Check whether title consists only of allowed characters
* Illegal characters are: / \ [ ] * ? : --> http://msdn.microsoft.com/en-us/library/ff837411.aspx
* Illegal characters not in documentation:   ' as first character
*--------------------------------------------------------------------*
    IF ip_title CA '/\[]*?:'.
      lv_errormessage = 'Found illegal character in sheetname. List of forbidden characters: /\[]*?:'(402).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

    IF ip_title IS NOT INITIAL AND ip_title(1) = `'`.
      lv_errormessage = 'Sheetname may not start with &'(403).   " & used instead of ' to allow fallbacklanguage
      REPLACE '&' IN lv_errormessage WITH `'`.
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.


*--------------------------------------------------------------------*
* Check whether title is unique in workbook
*--------------------------------------------------------------------*
    lo_worksheets_iterator = me->excel->get_worksheets_iterator( ).
    WHILE lo_worksheets_iterator->has_next( ) = 'X'.

      lo_worksheet ?= lo_worksheets_iterator->get_next( ).
      CHECK me->guid <> lo_worksheet->get_guid( ).  " Don't check against itself
      IF ip_title = lo_worksheet->get_title( ).  " Not unique --> raise exception
        errormessage = 'Duplicate sheetname &'.
        REPLACE '&' IN errormessage WITH ip_title.
        zcx_excel=>raise_text( errormessage ).
      ENDIF.

    ENDWHILE.

*--------------------------------------------------------------------*
* Remember old sheetname and rename sheet to desired name
*--------------------------------------------------------------------*
    CONCATENATE me->title '!' INTO lv_rangesheetname_old.
    me->title = ip_title.

*--------------------------------------------------------------------*
* After changing this worksheet's title we have to adjust
* all ranges that are referring to this worksheet.
*--------------------------------------------------------------------*
* 2do §1  -  Check if the following quickfix is solid
*           I fear it isn't - but this implementation is better then
*           nothing at all since it handles a supposed majority of cases
*--------------------------------------------------------------------*
    CONCATENATE me->title '!' INTO lv_rangesheetname_new.

    lo_ranges_iterator = me->excel->get_ranges_iterator( ).
    WHILE lo_ranges_iterator->has_next( ) = 'X'.

      lo_range ?= lo_ranges_iterator->get_next( ).
      lv_range_value = lo_range->get_value( ).
      REPLACE ALL OCCURRENCES OF lv_rangesheetname_old IN lv_range_value WITH lv_rangesheetname_new.
      IF sy-subrc = 0.
        lo_range->set_range_value( lv_range_value ).
      ENDIF.

    ENDWHILE.


  ENDMETHOD.                    "SET_TITLE


  METHOD update_dimension_range.

    DATA: ls_sheet_content TYPE zexcel_s_cell_data,
          lv_row_alpha     TYPE string,
          lv_column_alpha  TYPE zexcel_cell_column_alpha.

    CHECK sheet_content IS NOT INITIAL.

    upper_cell-cell_row = rows->get_min_index( ).
    IF upper_cell-cell_row = 0.
      upper_cell-cell_row = zcl_excel_common=>c_excel_sheet_max_row.
    ENDIF.
    upper_cell-cell_column = zcl_excel_common=>c_excel_sheet_max_col.

    lower_cell-cell_row = rows->get_max_index( ).
    IF lower_cell-cell_row = 0.
      lower_cell-cell_row = zcl_excel_common=>c_excel_sheet_min_row.
    ENDIF.
    lower_cell-cell_column = zcl_excel_common=>c_excel_sheet_min_col.

    LOOP AT sheet_content INTO ls_sheet_content.
      IF upper_cell-cell_row > ls_sheet_content-cell_row.
        upper_cell-cell_row = ls_sheet_content-cell_row.
      ENDIF.
      IF upper_cell-cell_column > ls_sheet_content-cell_column.
        upper_cell-cell_column = ls_sheet_content-cell_column.
      ENDIF.
      IF lower_cell-cell_row < ls_sheet_content-cell_row.
        lower_cell-cell_row = ls_sheet_content-cell_row.
      ENDIF.
      IF lower_cell-cell_column < ls_sheet_content-cell_column.
        lower_cell-cell_column = ls_sheet_content-cell_column.
      ENDIF.
    ENDLOOP.

    upper_cell-cell_coords = zcl_excel_common=>convert_column_a_row2columnrow( i_column = upper_cell-cell_column i_row = upper_cell-cell_row ).

    lower_cell-cell_coords = zcl_excel_common=>convert_column_a_row2columnrow( i_column = lower_cell-cell_column i_row = lower_cell-cell_row ).

  ENDMETHOD.                    "UPDATE_DIMENSION_RANGE


  METHOD zif_excel_sheet_printsettings~clear_print_repeat_columns.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
    CLEAR:  me->print_title_col_from,
            me->print_title_col_to  .


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
    me->print_title_set_range( ).


  ENDMETHOD.                    "ZIF_EXCEL_SHEET_PRINTSETTINGS~CLEAR_PRINT_REPEAT_COLUMNS


  METHOD zif_excel_sheet_printsettings~clear_print_repeat_rows.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
    CLEAR:  me->print_title_row_from,
            me->print_title_row_to  .


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
    me->print_title_set_range( ).


  ENDMETHOD.                    "ZIF_EXCEL_SHEET_PRINTSETTINGS~CLEAR_PRINT_REPEAT_ROWS


  METHOD zif_excel_sheet_printsettings~get_print_repeat_columns.
    ev_columns_from = me->print_title_col_from.
    ev_columns_to   = me->print_title_col_to.
  ENDMETHOD.                    "ZIF_EXCEL_SHEET_PRINTSETTINGS~GET_PRINT_REPEAT_COLUMNS


  METHOD zif_excel_sheet_printsettings~get_print_repeat_rows.
    ev_rows_from = me->print_title_row_from.
    ev_rows_to   = me->print_title_row_to.
  ENDMETHOD.                    "ZIF_EXCEL_SHEET_PRINTSETTINGS~GET_PRINT_REPEAT_ROWS


  METHOD zif_excel_sheet_printsettings~set_print_repeat_columns.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmöcker,                             2012-12-02
*--------------------------------------------------------------------*

    DATA: lv_col_from_int TYPE i,
          lv_col_to_int   TYPE i,
          lv_errormessage TYPE string.


    lv_col_from_int = zcl_excel_common=>convert_column2int( iv_columns_from ).
    lv_col_to_int   = zcl_excel_common=>convert_column2int( iv_columns_to ).

*--------------------------------------------------------------------*
* Check if valid range is supplied
*--------------------------------------------------------------------*
    IF lv_col_from_int < 1.
      lv_errormessage = 'Invalid range supplied for print-title repeatable columns'(401).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

    IF  lv_col_from_int > lv_col_to_int.
      lv_errormessage = 'Invalid range supplied for print-title repeatable columns'(401).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
    me->print_title_col_from = iv_columns_from.
    me->print_title_col_to   = iv_columns_to.


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
    me->print_title_set_range( ).

  ENDMETHOD.                    "ZIF_EXCEL_SHEET_PRINTSETTINGS~SET_PRINT_REPEAT_COLUMNS


  METHOD zif_excel_sheet_printsettings~set_print_repeat_rows.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmöcker,                             2012-12-02
*--------------------------------------------------------------------*

    DATA:     lv_errormessage                 TYPE string.


*--------------------------------------------------------------------*
* Check if valid range is supplied
*--------------------------------------------------------------------*
    IF iv_rows_from < 1.
      lv_errormessage = 'Invalid range supplied for print-title repeatable rowumns'(401).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

    IF  iv_rows_from > iv_rows_to.
      lv_errormessage = 'Invalid range supplied for print-title repeatable rowumns'(401).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

*--------------------------------------------------------------------*
* adjust internal representation
*--------------------------------------------------------------------*
    me->print_title_row_from = iv_rows_from.
    me->print_title_row_to   = iv_rows_to.


*--------------------------------------------------------------------*
* adjust corresponding range
*--------------------------------------------------------------------*
    me->print_title_set_range( ).


  ENDMETHOD.                    "ZIF_EXCEL_SHEET_PRINTSETTINGS~SET_PRINT_REPEAT_ROWS


  METHOD zif_excel_sheet_properties~get_right_to_left.
    result = right_to_left.
  ENDMETHOD.


  METHOD zif_excel_sheet_properties~get_style.
    IF zif_excel_sheet_properties~style IS NOT INITIAL.
      ep_style = zif_excel_sheet_properties~style.
    ELSE.
      ep_style = me->excel->get_default_style( ).
    ENDIF.
  ENDMETHOD.                    "ZIF_EXCEL_SHEET_PROPERTIES~GET_STYLE


  METHOD zif_excel_sheet_properties~initialize.

    zif_excel_sheet_properties~show_zeros   = zif_excel_sheet_properties=>c_showzero.
    zif_excel_sheet_properties~summarybelow = zif_excel_sheet_properties=>c_below_on.
    zif_excel_sheet_properties~summaryright = zif_excel_sheet_properties=>c_right_on.

* inizialize zoomscale values
    zif_excel_sheet_properties~zoomscale = 100.
    zif_excel_sheet_properties~zoomscale_normal = 100.
    zif_excel_sheet_properties~zoomscale_pagelayoutview = 100 .
    zif_excel_sheet_properties~zoomscale_sheetlayoutview = 100 .
  ENDMETHOD.                    "ZIF_EXCEL_SHEET_PROPERTIES~INITIALIZE


  METHOD zif_excel_sheet_properties~set_right_to_left.
    me->right_to_left = right_to_left.
  ENDMETHOD.


  METHOD zif_excel_sheet_properties~set_style.
    zif_excel_sheet_properties~style = ip_style.
  ENDMETHOD.                    "ZIF_EXCEL_SHEET_PROPERTIES~SET_STYLE


  METHOD zif_excel_sheet_protection~initialize.

    me->zif_excel_sheet_protection~protected = zif_excel_sheet_protection=>c_unprotected.
    CLEAR me->zif_excel_sheet_protection~password.
    me->zif_excel_sheet_protection~auto_filter            = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~delete_columns         = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~delete_rows            = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~format_cells           = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~format_columns         = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~format_rows            = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~insert_columns         = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~insert_hyperlinks      = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~insert_rows            = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~objects                = zif_excel_sheet_protection=>c_noactive.
*  me->zif_excel_sheet_protection~password               = zif_excel_sheet_protection=>c_noactive. "issue #68
    me->zif_excel_sheet_protection~pivot_tables           = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~protected              = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~scenarios              = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~select_locked_cells    = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~select_unlocked_cells  = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~sheet                  = zif_excel_sheet_protection=>c_noactive.
    me->zif_excel_sheet_protection~sort                   = zif_excel_sheet_protection=>c_noactive.

  ENDMETHOD.                    "ZIF_EXCEL_SHEET_PROTECTION~INITIALIZE


  METHOD zif_excel_sheet_vba_project~set_codename.
    me->zif_excel_sheet_vba_project~codename = ip_codename.
  ENDMETHOD.                    "ZIF_EXCEL_SHEET_VBA_PROJECT~SET_CODENAME


  METHOD zif_excel_sheet_vba_project~set_codename_pr.
    me->zif_excel_sheet_vba_project~codename_pr = ip_codename_pr.
  ENDMETHOD.                    "ZIF_EXCEL_SHEET_VBA_PROJECT~SET_CODENAME_PR
ENDCLASS.
