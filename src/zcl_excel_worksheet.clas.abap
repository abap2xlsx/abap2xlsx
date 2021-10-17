*----------------------------------------------------------------------*
*       CLASS ZCL_EXCEL_WORKSHEET DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
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
    TYPE-POOLS abap .
    TYPE-POOLS slis .
    TYPE-POOLS soi .

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
      END OF mty_s_ignored_errors,
      mty_th_ignored_errors TYPE HASHED TABLE OF mty_s_ignored_errors WITH UNIQUE KEY cell_coords,
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
    TYPES ty_doc_url TYPE c LENGTH 255.

    CONSTANTS c_break_column TYPE zexcel_break VALUE 2.     "#EC NOTEXT
    CONSTANTS c_break_none TYPE zexcel_break VALUE 0.       "#EC NOTEXT
    CONSTANTS c_break_row TYPE zexcel_break VALUE 1.        "#EC NOTEXT
    DATA excel TYPE REF TO zcl_excel READ-ONLY .
    DATA print_gridlines TYPE zexcel_print_gridlines READ-ONLY VALUE abap_false. "#EC NOTEXT
    DATA sheet_content TYPE zexcel_t_cell_data .
    DATA sheet_setup TYPE REF TO zcl_excel_sheet_setup .
    DATA show_gridlines TYPE zexcel_show_gridlines READ-ONLY VALUE abap_true. "#EC NOTEXT
    DATA show_rowcolheaders TYPE zexcel_show_gridlines READ-ONLY VALUE abap_true. "#EC NOTEXT
    DATA styles TYPE zexcel_t_sheet_style .
    DATA tabcolor TYPE zexcel_s_tabcolor READ-ONLY .
    DATA column_formulas TYPE mty_th_column_formula READ-ONLY .
    CLASS-DATA: BEGIN OF c_messages READ-ONLY,
                  formula_id_only_is_possible TYPE string,
                  column_formula_id_not_found TYPE string,
                  formula_not_in_this_table   TYPE string,
                  formula_in_other_column     TYPE string,
                END OF c_messages.

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
        VALUE(eo_column) TYPE REF TO zcl_excel_column .
    METHODS add_new_style_cond
      IMPORTING
        !ip_dimension_range  TYPE string DEFAULT 'A1'
      RETURNING
        VALUE(eo_style_cond) TYPE REF TO zcl_excel_style_cond.
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
      EXPORTING
        !es_table_settings      TYPE zexcel_s_table_settings
      RAISING
        zcx_excel .
    METHODS calculate_column_widths
      RAISING
        zcx_excel .
    METHODS change_cell_style
      IMPORTING
        !ip_column                      TYPE simple
        !ip_row                         TYPE zexcel_cell_row
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
    CLASS-METHODS class_constructor.
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
        !ip_column  TYPE simple
        !ip_row     TYPE zexcel_cell_row
      EXPORTING
        !ep_value   TYPE zexcel_cell_value
        !ep_rc      TYPE sysubrc
        !ep_style   TYPE REF TO zcl_excel_style
        !ep_guid    TYPE zexcel_cell_style
        !ep_formula TYPE zexcel_cell_formula
      RAISING
        zcx_excel .
    METHODS get_column
      IMPORTING
        !ip_column       TYPE simple
      RETURNING
        VALUE(eo_column) TYPE REF TO zcl_excel_column .
    METHODS get_columns
      RETURNING
        VALUE(eo_columns) TYPE REF TO zcl_excel_columns .
    METHODS get_columns_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO cl_object_collection_iterator .
    METHODS get_style_cond_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO cl_object_collection_iterator .
    METHODS get_data_validations_iterator
      RETURNING
        VALUE(eo_iterator) TYPE REF TO cl_object_collection_iterator .
    METHODS get_data_validations_size
      RETURNING
        VALUE(ep_size) TYPE i .
    METHODS get_default_column
      RETURNING
        VALUE(eo_column) TYPE REF TO zcl_excel_column .
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
        VALUE(eo_iterator) TYPE REF TO cl_object_collection_iterator .
    METHODS get_drawings_iterator
      IMPORTING
        !ip_type           TYPE zexcel_drawing_type
      RETURNING
        VALUE(eo_iterator) TYPE REF TO cl_object_collection_iterator .
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
        VALUE(eo_iterator) TYPE REF TO cl_object_collection_iterator .
    METHODS get_hyperlinks_size
      RETURNING
        VALUE(ep_size) TYPE i .
    METHODS get_ignored_errors
      RETURNING
        VALUE(rt_ignored_errors) TYPE mty_th_ignored_errors.
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
        VALUE(eo_iterator) TYPE REF TO cl_object_collection_iterator .
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
        VALUE(eo_iterator) TYPE REF TO cl_object_collection_iterator .
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
        VALUE(eo_iterator) TYPE REF TO cl_object_collection_iterator .
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
        !ip_column    TYPE simple
        !ip_row       TYPE zexcel_cell_row
        !ip_value     TYPE simple OPTIONAL
        !ip_formula   TYPE zexcel_cell_formula OPTIONAL
        !ip_style     TYPE zexcel_cell_style OPTIONAL
        !ip_hyperlink TYPE REF TO zcl_excel_hyperlink OPTIONAL
        !ip_data_type TYPE zexcel_cell_data_type OPTIONAL
        !ip_abap_type TYPE abap_typekind OPTIONAL
        !ip_column_formula_id TYPE mty_s_column_formula-id OPTIONAL
      RAISING
        zcx_excel .
    METHODS set_cell_formula
      IMPORTING
        !ip_column  TYPE simple
        !ip_row     TYPE zexcel_cell_row
        !ip_formula TYPE zexcel_cell_formula
      RAISING
        zcx_excel .
    METHODS set_cell_style
      IMPORTING
        !ip_column TYPE simple
        !ip_row    TYPE zexcel_cell_row
        !ip_style  TYPE zexcel_cell_style
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
        it_ignored_errors TYPE mty_th_ignored_errors.
    METHODS set_merge
      IMPORTING
        !ip_column_start TYPE simple DEFAULT zcl_excel_common=>c_excel_sheet_min_col
        !ip_column_end   TYPE simple DEFAULT zcl_excel_common=>c_excel_sheet_max_col
        !ip_row          TYPE zexcel_cell_row DEFAULT zcl_excel_common=>c_excel_sheet_min_row
        !ip_row_to       TYPE zexcel_cell_row DEFAULT zcl_excel_common=>c_excel_sheet_max_row
        !ip_style        TYPE zexcel_cell_style OPTIONAL "added parameter
        !ip_value        TYPE simple OPTIONAL "added parameter
        !ip_formula      TYPE zexcel_cell_formula OPTIONAL "added parameter
      RAISING
        zcx_excel .
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
        !ip_hdr_style       TYPE zexcel_cell_style OPTIONAL
        !ip_body_style      TYPE zexcel_cell_style OPTIONAL
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
        !iv_skipped_rows TYPE int4 DEFAULT 0
        !iv_skipped_cols TYPE int4 DEFAULT 0
        !iv_max_col      TYPE int4 OPTIONAL
        !iv_max_row      TYPE int4 OPTIONAL
      EXPORTING
        !et_table        TYPE STANDARD TABLE
      RAISING
        zcx_excel .
    METHODS set_merge_style
      IMPORTING
        !ip_column_start TYPE simple OPTIONAL
        !ip_column_end   TYPE simple OPTIONAL
        !ip_row          TYPE zexcel_cell_row OPTIONAL
        !ip_row_to       TYPE zexcel_cell_row OPTIONAL
        !ip_style        TYPE zexcel_cell_style OPTIONAL .
    METHODS set_area_formula
      IMPORTING
        !ip_column_start TYPE simple
        !ip_column_end   TYPE simple OPTIONAL
        !ip_row          TYPE zexcel_cell_row
        !ip_row_to       TYPE zexcel_cell_row OPTIONAL
        !ip_formula      TYPE zexcel_cell_formula
        !ip_merge        TYPE abap_bool OPTIONAL
      RAISING
        zcx_excel .
    METHODS set_area_style
      IMPORTING
        !ip_column_start TYPE simple
        !ip_column_end   TYPE simple OPTIONAL
        !ip_row          TYPE zexcel_cell_row
        !ip_row_to       TYPE zexcel_cell_row OPTIONAL
        !ip_style        TYPE zexcel_cell_style
        !ip_merge        TYPE abap_bool OPTIONAL .
    METHODS set_area
      IMPORTING
        !ip_column_start TYPE simple
        !ip_column_end   TYPE simple OPTIONAL
        !ip_row          TYPE zexcel_cell_row
        !ip_row_to       TYPE zexcel_cell_row OPTIONAL
        !ip_value        TYPE simple OPTIONAL
        !ip_formula      TYPE zexcel_cell_formula OPTIONAL
        !ip_style        TYPE zexcel_cell_style OPTIONAL
        !ip_hyperlink    TYPE REF TO zcl_excel_hyperlink OPTIONAL
        !ip_data_type    TYPE zexcel_cell_data_type OPTIONAL
        !ip_abap_type    TYPE abap_typekind OPTIONAL
        !ip_merge        TYPE abap_bool OPTIONAL
      RAISING
        zcx_excel .
    METHODS get_header_footer_drawings
      RETURNING
        VALUE(rt_drawings) TYPE zexcel_t_drawings .
    METHODS set_area_hyperlink
      IMPORTING
        !ip_column_start TYPE simple
        !ip_column_end   TYPE simple OPTIONAL
        !ip_row          TYPE zexcel_cell_row
        !ip_row_to       TYPE zexcel_cell_row OPTIONAL
        !ip_url          TYPE string
        !ip_is_internal  TYPE abap_bool
      RAISING
        zcx_excel .
  PROTECTED SECTION.
  PRIVATE SECTION.

    TYPES:
      BEGIN OF mty_s_font_metric,
        char       TYPE c LENGTH 1,
        char_width TYPE tdcwidths,
      END OF mty_s_font_metric .
    TYPES:
      mty_th_font_metrics
             TYPE HASHED TABLE OF mty_s_font_metric
             WITH UNIQUE KEY char .
    TYPES:
      BEGIN OF mty_s_font_cache,
        font_name       TYPE zexcel_style_font_name,
        font_height     TYPE tdfontsize,
        flag_bold       TYPE abap_bool,
        flag_italic     TYPE abap_bool,
        th_font_metrics TYPE mty_th_font_metrics,
      END OF mty_s_font_cache .
    TYPES:
      mty_th_font_cache
             TYPE HASHED TABLE OF mty_s_font_cache
             WITH UNIQUE KEY font_name font_height flag_bold flag_italic .
    TYPES:
*  types:
*    mty_ts_row_dimension TYPE SORTED TABLE OF zexcel_s_worksheet_rowdimensio WITH UNIQUE KEY row .
      BEGIN OF mty_merge,
        row_from TYPE i,
        row_to   TYPE i,
        col_from TYPE i,
        col_to   TYPE i,
      END OF mty_merge .
    TYPES:
      mty_ts_merge TYPE SORTED TABLE OF mty_merge WITH UNIQUE KEY table_line .

*"* private components of class ZCL_EXCEL_WORKSHEET
*"* do not include other source files here!!!
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
    DATA hyperlinks TYPE REF TO cl_object_collection .
    DATA lower_cell TYPE zexcel_s_cell_data .
    DATA mo_pagebreaks TYPE REF TO zcl_excel_worksheet_pagebreaks .
    CLASS-DATA mth_font_cache TYPE mty_th_font_cache .
    DATA mt_merged_cells TYPE mty_ts_merge .
    DATA mt_row_outlines TYPE mty_ts_outlines_row .
    DATA print_title_col_from TYPE zexcel_cell_column_alpha .
    DATA print_title_col_to TYPE zexcel_cell_column_alpha .
    DATA print_title_row_from TYPE zexcel_cell_row .
    DATA print_title_row_to TYPE zexcel_cell_row .
    DATA ranges TYPE REF TO zcl_excel_ranges .
    DATA rows TYPE REF TO zcl_excel_rows .
    DATA tables TYPE REF TO cl_object_collection .
    DATA title TYPE zexcel_sheet_title VALUE 'Worksheet'. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  . " .
    DATA upper_cell TYPE zexcel_s_cell_data .
    DATA mt_ignored_errors TYPE mty_th_ignored_errors.

    METHODS calculate_cell_width
      IMPORTING
        !ip_column      TYPE simple
        !ip_row         TYPE zexcel_cell_row
      RETURNING
        VALUE(ep_width) TYPE float
      RAISING
        zcx_excel .
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
    METHODS generate_title
      RETURNING
        VALUE(ep_title) TYPE zexcel_sheet_title .
    METHODS get_value_type
      IMPORTING
        !ip_value      TYPE simple
      EXPORTING
        !ep_value      TYPE simple
        !ep_value_type TYPE abap_typekind .
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
        others                  = 8.
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
      lc_top_left_row    TYPE zexcel_cell_row VALUE 1.

    DATA:
      lv_row_int            TYPE zexcel_cell_row,
      lv_first_row          TYPE zexcel_cell_row,
      lv_last_row           TYPE zexcel_cell_row,
      lv_column_int         TYPE zexcel_cell_column,
      lv_column_alpha       TYPE zexcel_cell_column_alpha,
      lt_field_catalog      TYPE zexcel_t_fieldcatalog,
      lv_id                 TYPE i,
      lv_rows               TYPE i,
      lv_formula            TYPE string,
      ls_settings           TYPE zexcel_s_table_settings,
      lo_table              TYPE REF TO zcl_excel_table,
      lt_column_name_buffer TYPE SORTED TABLE OF string WITH UNIQUE KEY table_line,
      lv_value              TYPE string,
      lv_value_lowercase    TYPE string,
      lv_syindex            TYPE c LENGTH 3,
      lv_errormessage       TYPE string,                                            "ins issue #237

      lv_columns            TYPE i,
      lt_columns            TYPE zexcel_t_fieldcatalog,
      lv_maxcol             TYPE i,
      lv_maxrow             TYPE i,
      lo_iterator           TYPE REF TO cl_object_collection_iterator,
      lo_style_cond         TYPE REF TO zcl_excel_style_cond,
      lo_curtable           TYPE REF TO zcl_excel_table.
    DATA: ls_column_formula TYPE mty_s_column_formula,
          lv_mincol         TYPE i.

    FIELD-SYMBOLS:
      <ls_field_catalog>        TYPE zexcel_s_fieldcatalog,
      <ls_field_catalog_custom> TYPE zexcel_s_fieldcatalog,
      <fs_table_line>           TYPE any,
      <fs_fldval>               TYPE any.

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
      lt_field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = ip_table ).
    ELSE.
      lt_field_catalog = it_field_catalog.
    ENDIF.

    SORT lt_field_catalog BY position.

*--------------------------------------------------------------------*
*  issue #237   Check if overlapping areas exist  Start
*--------------------------------------------------------------------*
    "Get the number of columns for the current table
    lt_columns = lt_field_catalog.
    DELETE lt_columns WHERE dynpfld NE abap_true.
    DESCRIBE TABLE lt_columns LINES lv_columns.

    "Calculate the top left row of the current table
    lv_column_int = zcl_excel_common=>convert_column2int( ls_settings-top_left_column ).
    lv_row_int    = ls_settings-top_left_row.

    "Get number of row for the current table
    DESCRIBE TABLE ip_table LINES lv_rows.

    "Calculate the bottom right row for the current table
    lv_maxcol                       = lv_column_int + lv_columns - 1.
    lv_maxrow                       = lv_row_int    + lv_rows - 1.
    ls_settings-bottom_right_column = zcl_excel_common=>convert_column2alpha( lv_maxcol ).
    ls_settings-bottom_right_row    = lv_maxrow.

    lv_column_int                   = zcl_excel_common=>convert_column2int( ls_settings-top_left_column ).

    lo_iterator = me->tables->get_iterator( ).
    WHILE lo_iterator->has_next( ) EQ abap_true.

      lo_curtable ?= lo_iterator->get_next( ).
      IF  (    (  ls_settings-top_left_row     GE lo_curtable->settings-top_left_row                             AND ls_settings-top_left_row     LE lo_curtable->settings-bottom_right_row )
            OR
               (  ls_settings-bottom_right_row GE lo_curtable->settings-top_left_row                             AND ls_settings-bottom_right_row LE lo_curtable->settings-bottom_right_row )
          )
        AND
          (    (  lv_column_int GE zcl_excel_common=>convert_column2int( lo_curtable->settings-top_left_column ) AND lv_column_int LE zcl_excel_common=>convert_column2int( lo_curtable->settings-bottom_right_column ) )
            OR
               (  lv_maxcol     GE zcl_excel_common=>convert_column2int( lo_curtable->settings-top_left_column ) AND lv_maxcol     LE zcl_excel_common=>convert_column2int( lo_curtable->settings-bottom_right_column ) )
          ).
        lv_errormessage = 'Table overlaps with previously bound table and will not be added to worksheet.'(400).
        zcx_excel=>raise_text( lv_errormessage ).
      ENDIF.

    ENDWHILE.
*--------------------------------------------------------------------*
*  issue #237   Check if overlapping areas exist  End
*--------------------------------------------------------------------*

    CREATE OBJECT lo_table.
    lo_table->settings = ls_settings.
    lo_table->set_data( ir_data = ip_table ).
    lv_id = me->excel->get_next_table_id( ).
    lo_table->set_id( iv_id = lv_id ).
*  lo_table->fieldcat = lt_field_catalog[].

    me->tables->add( lo_table ).

* It is better to loop column by column (only visible column)
    LOOP AT lt_field_catalog ASSIGNING <ls_field_catalog> WHERE dynpfld EQ abap_true.

      lv_column_alpha = zcl_excel_common=>convert_column2alpha( lv_column_int ).

      " Due restrinction of new table object we cannot have two column with the same name
      " Check if a column with the same name exists, if exists add a counter
      " If no medium description is provided we try to use small or long
*    lv_value = <ls_field_catalog>-scrtext_m.
      FIELD-SYMBOLS: <scrtxt1> TYPE any,
                     <scrtxt2> TYPE any,
                     <scrtxt3> TYPE any.

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
        lv_value = <scrtxt1>.
        <ls_field_catalog>-scrtext_l = lv_value.
      ELSEIF <scrtxt2> IS NOT INITIAL.
        lv_value = <scrtxt2>.
        <ls_field_catalog>-scrtext_l = lv_value.
      ELSEIF <scrtxt3> IS NOT INITIAL.
        lv_value = <scrtxt3>.
        <ls_field_catalog>-scrtext_l = lv_value.
      ELSE.
        lv_value = 'Column'.  " default value as Excel does
        <ls_field_catalog>-scrtext_l = lv_value.
      ENDIF.
      WHILE 1 = 1.
        lv_value_lowercase = lv_value.
        TRANSLATE lv_value_lowercase TO LOWER CASE.
        READ TABLE lt_column_name_buffer TRANSPORTING NO FIELDS WITH KEY table_line = lv_value_lowercase BINARY SEARCH.
        IF sy-subrc <> 0.
          <ls_field_catalog>-scrtext_l = lv_value.
          INSERT lv_value_lowercase INTO TABLE lt_column_name_buffer.
          EXIT.
        ELSE.
          lv_syindex = sy-index.
          CONCATENATE <ls_field_catalog>-scrtext_l lv_syindex INTO lv_value.
        ENDIF.

      ENDWHILE.
      " First of all write column header
      IF <ls_field_catalog>-style_header IS NOT INITIAL.
        me->set_cell( ip_column = lv_column_alpha
                      ip_row    = lv_row_int
                      ip_value  = lv_value
                      ip_style  = <ls_field_catalog>-style_header ).
      ELSE.
        me->set_cell( ip_column = lv_column_alpha
                      ip_row    = lv_row_int
                      ip_value  = lv_value ).
      ENDIF.

      IF <ls_field_catalog>-column_formula IS NOT INITIAL.
        ls_column_formula-id                     = lines( column_formulas ) + 1.
        ls_column_formula-column                 = lv_column_int.
        ls_column_formula-formula                = <ls_field_catalog>-column_formula.
        ls_column_formula-table_top_left_row     = lo_table->settings-top_left_row.
        ls_column_formula-table_bottom_right_row = lo_table->settings-bottom_right_row.
        ls_column_formula-table_left_column_int  = lv_mincol.
        ls_column_formula-table_right_column_int = lv_maxcol.
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
          IF <ls_field_catalog>-style IS NOT INITIAL.
            IF <ls_field_catalog>-abap_type IS NOT INITIAL.
              me->set_cell( ip_column = lv_column_alpha
                          ip_row    = lv_row_int
                          ip_value  = <fs_fldval>
                          ip_abap_type = <ls_field_catalog>-abap_type
                          ip_style  = <ls_field_catalog>-style ).
            ELSE.
              me->set_cell( ip_column = lv_column_alpha
                            ip_row    = lv_row_int
                            ip_value  = <fs_fldval>
                            ip_style  = <ls_field_catalog>-style ).
            ENDIF.
          ELSE.
            IF <ls_field_catalog>-abap_type IS NOT INITIAL.
              me->set_cell( ip_column = lv_column_alpha
                          ip_row    = lv_row_int
                          ip_abap_type = <ls_field_catalog>-abap_type
                          ip_value  = <fs_fldval> ).
            ELSE.
              me->set_cell( ip_column = lv_column_alpha
                            ip_row    = lv_row_int
                            ip_value  = <fs_fldval> ).
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
        lv_formula = lo_table->get_totals_formula( ip_column = <ls_field_catalog>-scrtext_l ip_function = <ls_field_catalog>-totals_function ).
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
        lv_last_row     = ls_settings-top_left_row + lv_rows.
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
      es_table_settings-bottom_right_row    = ls_settings-top_left_row + lv_rows + 1. "Last rows
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

    CONSTANTS:
      lc_default_font_name   TYPE zexcel_style_font_name VALUE 'Calibri', "#EC NOTEXT
      lc_default_font_height TYPE tdfontsize VALUE '110',
      lc_excel_cell_padding  TYPE float VALUE '0.75'.

    DATA: ld_cell_value                TYPE zexcel_cell_value,
          ld_current_character         TYPE c LENGTH 1,
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
          ld_font_height               TYPE tdfontsize VALUE lc_default_font_height,
          lt_itcfc                     TYPE STANDARD TABLE OF itcfc,
          ld_offset                    TYPE i,
          ld_length                    TYPE i,
          ld_uccp                      TYPE i,
          ls_font_metric               TYPE mty_s_font_metric,
          ld_width_from_font_metrics   TYPE i,
          ld_font_family               TYPE itcfh-tdfamily,
          ld_font_name                 TYPE zexcel_style_font_name VALUE lc_default_font_name,
          lt_font_families             LIKE STANDARD TABLE OF ld_font_family,
          ls_font_cache                TYPE mty_s_font_cache.

    FIELD-SYMBOLS: <ls_font_cache>  TYPE mty_s_font_cache,
                   <ls_font_metric> TYPE mty_s_font_metric,
                   <ls_itcfc>       TYPE itcfc.

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

    " Check if the same font (font name and font attributes) was already
    " used before
    READ TABLE mth_font_cache
      WITH TABLE KEY
        font_name   = ld_font_name
        font_height = ld_font_height
        flag_bold   = ld_flag_bold
        flag_italic = ld_flag_italic
      ASSIGNING <ls_font_cache>.

    IF sy-subrc <> 0.
      " Font is used for the first time
      " Add the font to our local font cache
      ls_font_cache-font_name   = ld_font_name.
      ls_font_cache-font_height = ld_font_height.
      ls_font_cache-flag_bold   = ld_flag_bold.
      ls_font_cache-flag_italic = ld_flag_italic.
      INSERT ls_font_cache INTO TABLE mth_font_cache
        ASSIGNING <ls_font_cache>.

      " Determine the SAPscript font family name from the Excel
      " font name
      SELECT tdfamily
        FROM tfo01
        INTO TABLE lt_font_families
        UP TO 1 ROWS
        WHERE tdtext = ld_font_name
        ORDER BY PRIMARY KEY.

      " Check if a matching font family was found
      " Fonts can be uploaded from TTF files using transaction SE73
      IF lines( lt_font_families ) > 0.
        READ TABLE lt_font_families INDEX 1 INTO ld_font_family.

        " Load font metrics (returns a table with the size of each letter
        " in the font)
        CALL FUNCTION 'LOAD_FONT'
          EXPORTING
            family      = ld_font_family
            height      = ld_font_height
            printer     = 'SWIN'
            bold        = ld_flag_bold
            italic      = ld_flag_italic
          TABLES
            metric      = lt_itcfc
          EXCEPTIONS
            font_family = 1
            codepage    = 2
            device_type = 3
            OTHERS      = 4.
        IF sy-subrc <> 0.
          CLEAR lt_itcfc.
        ENDIF.

        " For faster access, convert each character number to the actual
        " character, and store the characters and their sizes in a hash
        " table
        LOOP AT lt_itcfc ASSIGNING <ls_itcfc>.
          ld_uccp = <ls_itcfc>-cpcharno.
          ls_font_metric-char =
            cl_abap_conv_in_ce=>uccpi( ld_uccp ).
          ls_font_metric-char_width = <ls_itcfc>-tdcwidths.
          INSERT ls_font_metric
            INTO TABLE <ls_font_cache>-th_font_metrics.
        ENDLOOP.

      ENDIF.
    ENDIF.

    " Calculate the cell width
    " If available, use font metrics
    IF lines( <ls_font_cache>-th_font_metrics ) = 0.
      " Font metrics are not available
      " -> Calculate the cell width using only the font size
      ld_length = strlen( ld_cell_value ).
      ep_width = ld_length * ld_font_height / lc_default_font_height + lc_excel_cell_padding.

    ELSE.
      " Font metrics are available

      " Calculate the size of the text by adding the sizes of each
      " letter
      ld_length = strlen( ld_cell_value ).
      DO ld_length TIMES.
        " Subtract 1, because the first character is at offset 0
        ld_offset = sy-index - 1.

        " Read the current character from the cell value
        ld_current_character = ld_cell_value+ld_offset(1).

        " Look up the size of the current letter
        READ TABLE <ls_font_cache>-th_font_metrics
          WITH TABLE KEY char = ld_current_character
          ASSIGNING <ls_font_metric>.
        IF sy-subrc = 0.
          " The size of the letter is known
          " -> Add the actual size of the letter
          ADD <ls_font_metric>-char_width TO ld_width_from_font_metrics.
        ELSE.
          " The size of the letter is unknown
          " -> Add the font height as the default letter size
          ADD ld_font_height TO ld_width_from_font_metrics.
        ENDIF.
      ENDDO.

      " Add cell padding (Excel makes columns a bit wider than the space
      " that is needed for the text itself) and convert unit
      " (division by 100)
      ep_width = ld_width_from_font_metrics / 100 + lc_excel_cell_padding.
    ENDIF.

    " If the current cell contains an auto filter, make it a bit wider.
    " The size used by the auto filter button does not depend on the font
    " size.
    IF ld_flag_contains_auto_filter = abap_true.
      ADD 2 TO ep_width.
    ENDIF.

  ENDMETHOD.                    "CALCULATE_CELL_WIDTH


  METHOD calculate_column_widths.
    TYPES:
      BEGIN OF t_auto_size,
        col_index TYPE int4,
        width     TYPE float,
      END   OF t_auto_size.
    TYPES: tt_auto_size TYPE TABLE OF t_auto_size.

    DATA: lo_column_iterator TYPE REF TO cl_object_collection_iterator,
          lo_column          TYPE REF TO zcl_excel_column.

    DATA: auto_size   TYPE flag.
    DATA: auto_sizes  TYPE tt_auto_size.
    DATA: count       TYPE int4.
    DATA: highest_row TYPE int4.
    DATA: width       TYPE float.

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


  METHOD change_cell_style.
    " issue # 139
    DATA: stylemapping    TYPE zexcel_s_stylemapping,

          complete_style  TYPE zexcel_s_cstyle_complete,
          complete_stylex TYPE zexcel_s_cstylex_complete,

          borderx         TYPE zexcel_s_cstylex_border,
          l_guid          TYPE zexcel_cell_style.   "issue # 177

* We have a lot of parameters.  Use some macros to make the coding more structured

    DEFINE clear_initial_colorxfields.
      IF &1-rgb IS INITIAL.
        CLEAR &2-rgb.
      ENDIF.
      IF &1-indexed IS INITIAL.
        CLEAR &2-indexed.
      ENDIF.
      IF &1-theme IS INITIAL.
        CLEAR &2-theme.
      ENDIF.
      IF &1-tint IS INITIAL.
        CLEAR &2-tint.
      ENDIF.
    END-OF-DEFINITION.

    DEFINE move_supplied_borders.
      IF ip_&1 IS SUPPLIED.  " only act if parameter was supplied
        IF ip_x&1 IS SUPPLIED.  "
          borderx = ip_x&1.          " use supplied x-parameter
        ELSE.
          CLEAR borderx WITH 'X'.
* clear in a way that would be expected to work easily
          IF ip_&1-border_style IS  INITIAL.
            CLEAR borderx-border_style.
          ENDIF.
          clear_initial_colorxfields ip_&1-border_color borderx-border_color.
        ENDIF.
        MOVE-CORRESPONDING ip_&1   TO complete_style-&2.
        MOVE-CORRESPONDING borderx TO complete_stylex-&2.
      ENDIF.
    END-OF-DEFINITION.

* First get current stylsettings
    TRY.
        me->get_cell( EXPORTING ip_column = ip_column  " Cell Column
                                ip_row    = ip_row      " Cell Row
                      IMPORTING ep_guid   = l_guid )." Cell Value ).  "issue # 177


        stylemapping = me->excel->get_style_to_guid( l_guid ).        "issue # 177
        complete_style  = stylemapping-complete_style.
        complete_stylex = stylemapping-complete_stylex.
      CATCH zcx_excel.
* Error --> use submitted style
    ENDTRY.

*  move_supplied_multistyles: complete.
    IF ip_complete IS SUPPLIED.
      IF ip_xcomplete IS NOT SUPPLIED.
        zcx_excel=>raise_text( 'Complete styleinfo has to be supplied with corresponding X-field' ).
      ENDIF.
      MOVE-CORRESPONDING ip_complete  TO complete_style.
      MOVE-CORRESPONDING ip_xcomplete TO complete_stylex.
    ENDIF.



    IF ip_font IS SUPPLIED.
      DATA: fontx LIKE ip_xfont.
      IF ip_xfont IS SUPPLIED.
        fontx = ip_xfont.
      ELSE.
* Only supplied values should be used - exception: Flags bold and italic strikethrough underline
        MOVE 'X' TO: fontx-bold,
                     fontx-italic,
                     fontx-strikethrough,
                     fontx-underline_mode.
        CLEAR fontx-color WITH 'X'.
        clear_initial_colorxfields ip_font-color fontx-color.
        IF ip_font-family IS NOT INITIAL.
          fontx-family = 'X'.
        ENDIF.
        IF ip_font-name IS NOT INITIAL.
          fontx-name = 'X'.
        ENDIF.
        IF ip_font-scheme IS NOT INITIAL.
          fontx-scheme = 'X'.
        ENDIF.
        IF ip_font-size IS NOT INITIAL.
          fontx-size = 'X'.
        ENDIF.
        IF ip_font-underline_mode IS NOT INITIAL.
          fontx-underline_mode = 'X'.
        ENDIF.
      ENDIF.
      MOVE-CORRESPONDING ip_font  TO complete_style-font.
      MOVE-CORRESPONDING fontx    TO complete_stylex-font.
* Correction for undeline mode
    ENDIF.

    IF ip_fill IS SUPPLIED.
      DATA: fillx LIKE ip_xfill.
      IF ip_xfill IS SUPPLIED.
        fillx = ip_xfill.
      ELSE.
        CLEAR fillx WITH 'X'.
        IF ip_fill-filltype IS INITIAL.
          CLEAR fillx-filltype.
        ENDIF.
        clear_initial_colorxfields ip_fill-fgcolor fillx-fgcolor.
        clear_initial_colorxfields ip_fill-bgcolor fillx-bgcolor.

      ENDIF.
      MOVE-CORRESPONDING ip_fill  TO complete_style-fill.
      MOVE-CORRESPONDING fillx    TO complete_stylex-fill.
    ENDIF.


    IF ip_borders IS SUPPLIED.
      DATA: bordersx LIKE ip_xborders.
      IF ip_xborders IS SUPPLIED.
        bordersx = ip_xborders.
      ELSE.
        CLEAR bordersx WITH 'X'.
        IF ip_borders-allborders-border_style IS INITIAL.
          CLEAR bordersx-allborders-border_style.
        ENDIF.
        IF ip_borders-diagonal-border_style IS INITIAL.
          CLEAR bordersx-diagonal-border_style.
        ENDIF.
        IF ip_borders-down-border_style IS INITIAL.
          CLEAR bordersx-down-border_style.
        ENDIF.
        IF ip_borders-left-border_style IS INITIAL.
          CLEAR bordersx-left-border_style.
        ENDIF.
        IF ip_borders-right-border_style IS INITIAL.
          CLEAR bordersx-right-border_style.
        ENDIF.
        IF ip_borders-top-border_style IS INITIAL.
          CLEAR bordersx-top-border_style.
        ENDIF.
        clear_initial_colorxfields ip_borders-allborders-border_color bordersx-allborders-border_color.
        clear_initial_colorxfields ip_borders-diagonal-border_color   bordersx-diagonal-border_color.
        clear_initial_colorxfields ip_borders-down-border_color       bordersx-down-border_color.
        clear_initial_colorxfields ip_borders-left-border_color       bordersx-left-border_color.
        clear_initial_colorxfields ip_borders-right-border_color      bordersx-right-border_color.
        clear_initial_colorxfields ip_borders-top-border_color        bordersx-top-border_color.

      ENDIF.
      MOVE-CORRESPONDING ip_borders  TO complete_style-borders.
      MOVE-CORRESPONDING bordersx    TO complete_stylex-borders.
    ENDIF.

    IF ip_alignment IS SUPPLIED.
      DATA: alignmentx LIKE ip_xalignment.
      IF ip_xalignment IS SUPPLIED.
        alignmentx = ip_xalignment.
      ELSE.
        CLEAR alignmentx WITH 'X'.
        IF ip_alignment-horizontal IS INITIAL.
          CLEAR alignmentx-horizontal.
        ENDIF.
        IF ip_alignment-vertical IS INITIAL.
          CLEAR alignmentx-vertical.
        ENDIF.
      ENDIF.
      MOVE-CORRESPONDING ip_alignment  TO complete_style-alignment.
      MOVE-CORRESPONDING alignmentx    TO complete_stylex-alignment.
    ENDIF.

    IF ip_protection IS SUPPLIED.
      MOVE-CORRESPONDING ip_protection  TO complete_style-protection.
      IF ip_xprotection IS SUPPLIED.
        MOVE-CORRESPONDING ip_xprotection TO complete_stylex-protection.
      ELSE.
        IF ip_protection-hidden IS NOT INITIAL.
          complete_stylex-protection-hidden = 'X'.
        ENDIF.
        IF ip_protection-locked IS NOT INITIAL.
          complete_stylex-protection-locked = 'X'.
        ENDIF.
      ENDIF.
    ENDIF.


    move_supplied_borders    : borders_allborders borders-allborders,
                               borders_diagonal   borders-diagonal  ,
                               borders_down       borders-down      ,
                               borders_left       borders-left      ,
                               borders_right      borders-right     ,
                               borders_top        borders-top       .

    DEFINE move_supplied_singlestyles.
      IF ip_&1 IS SUPPLIED.
        complete_style-&2 = ip_&1.
        complete_stylex-&2 = 'X'.
      ENDIF.
    END-OF-DEFINITION.

    move_supplied_singlestyles: number_format_format_code  number_format-format_code,
                                font_bold                     font-bold,
                                font_color                    font-color,
                                font_color_rgb                font-color-rgb,
                                font_color_indexed            font-color-indexed,
                                font_color_theme              font-color-theme,
                                font_color_tint               font-color-tint,

                                font_family                   font-family,
                                font_italic                   font-italic,
                                font_name                     font-name,
                                font_scheme                   font-scheme,
                                font_size                     font-size,
                                font_strikethrough            font-strikethrough,
                                font_underline                font-underline,
                                font_underline_mode           font-underline_mode,
                                fill_filltype                 fill-filltype,
                                fill_rotation                 fill-rotation,
                                fill_fgcolor                  fill-fgcolor,
                                fill_fgcolor_rgb              fill-fgcolor-rgb,
                                fill_fgcolor_indexed          fill-fgcolor-indexed,
                                fill_fgcolor_theme            fill-fgcolor-theme,
                                fill_fgcolor_tint             fill-fgcolor-tint,

                                fill_bgcolor                  fill-bgcolor,
                                fill_bgcolor_rgb              fill-bgcolor-rgb,
                                fill_bgcolor_indexed          fill-bgcolor-indexed,
                                fill_bgcolor_theme            fill-bgcolor-theme,
                                fill_bgcolor_tint             fill-bgcolor-tint,

                                fill_gradtype_type            fill-gradtype-type,
                                fill_gradtype_degree          fill-gradtype-degree,
                                fill_gradtype_bottom          fill-gradtype-bottom,
                                fill_gradtype_left            fill-gradtype-left,
                                fill_gradtype_top             fill-gradtype-top,
                                fill_gradtype_right           fill-gradtype-right,
                                fill_gradtype_position1       fill-gradtype-position1,
                                fill_gradtype_position2       fill-gradtype-position2,
                                fill_gradtype_position3       fill-gradtype-position3,



                                borders_diagonal_mode         borders-diagonal_mode,
                                alignment_horizontal          alignment-horizontal,
                                alignment_vertical            alignment-vertical,
                                alignment_textrotation        alignment-textrotation,
                                alignment_wraptext            alignment-wraptext,
                                alignment_shrinktofit         alignment-shrinktofit,
                                alignment_indent              alignment-indent,
                                protection_hidden             protection-hidden,
                                protection_locked             protection-locked,

                                borders_allborders_style      borders-allborders-border_style,
                                borders_allborders_color      borders-allborders-border_color,
                                borders_allbo_color_rgb       borders-allborders-border_color-rgb,
                                borders_allbo_color_indexed   borders-allborders-border_color-indexed,
                                borders_allbo_color_theme     borders-allborders-border_color-theme,
                                borders_allbo_color_tint      borders-allborders-border_color-tint,

                                borders_diagonal_style        borders-diagonal-border_style,
                                borders_diagonal_color        borders-diagonal-border_color,
                                borders_diagonal_color_rgb    borders-diagonal-border_color-rgb,
                                borders_diagonal_color_inde   borders-diagonal-border_color-indexed,
                                borders_diagonal_color_them   borders-diagonal-border_color-theme,
                                borders_diagonal_color_tint   borders-diagonal-border_color-tint,

                                borders_down_style            borders-down-border_style,
                                borders_down_color            borders-down-border_color,
                                borders_down_color_rgb        borders-down-border_color-rgb,
                                borders_down_color_indexed    borders-down-border_color-indexed,
                                borders_down_color_theme      borders-down-border_color-theme,
                                borders_down_color_tint       borders-down-border_color-tint,

                                borders_left_style            borders-left-border_style,
                                borders_left_color            borders-left-border_color,
                                borders_left_color_rgb        borders-left-border_color-rgb,
                                borders_left_color_indexed    borders-left-border_color-indexed,
                                borders_left_color_theme      borders-left-border_color-theme,
                                borders_left_color_tint       borders-left-border_color-tint,

                                borders_right_style           borders-right-border_style,
                                borders_right_color           borders-right-border_color,
                                borders_right_color_rgb       borders-right-border_color-rgb,
                                borders_right_color_indexed   borders-right-border_color-indexed,
                                borders_right_color_theme     borders-right-border_color-theme,
                                borders_right_color_tint      borders-right-border_color-tint,

                                borders_top_style             borders-top-border_style,
                                borders_top_color             borders-top-border_color,
                                borders_top_color_rgb         borders-top-border_color-rgb,
                                borders_top_color_indexed     borders-top-border_color-indexed,
                                borders_top_color_theme       borders-top-border_color-theme,
                                borders_top_color_tint        borders-top-border_color-tint.


* Now we have a completly filled styles.
* This can be used to get the guid
* Return guid if requested.  Might be used if copy&paste of styles is requested
    ep_guid = me->excel->get_static_cellstyle_guid( ip_cstyle_complete  = complete_style
                                                    ip_cstylex_complete = complete_stylex  ).
    me->set_cell_style( ip_column = ip_column
                        ip_row    = ip_row
                        ip_style  = ep_guid ).

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


  METHOD class_constructor.

    c_messages-formula_id_only_is_possible = |{ 'If Formula ID is used, value and formula must be empty'(008) }|.
    c_messages-column_formula_id_not_found = |{ 'The Column Formula does not exist'(009) }|.
    c_messages-formula_not_in_this_table = |{ 'The cell uses a Column Formula which should be part of the same table'(010) }|.
    c_messages-formula_in_other_column = |{ 'The cell uses a Column Formula which is in a different column'(011) }|.

  ENDMETHOD.


  METHOD constructor.
    DATA: lv_title TYPE zexcel_sheet_title.

    me->excel = ip_excel.

*  CALL FUNCTION 'GUID_CREATE'                                    " del issue #379 - function is outdated in newer releases
*    IMPORTING
*      ev_guid_16 = me->guid.
    me->guid = zcl_excel_obsolete_func_wrap=>guid_create( ).        " ins issue #379 - replacement for outdated function call

    IF ip_title IS NOT INITIAL.
      lv_title = ip_title.
    ELSE.
*    lv_title = me->guid.             " del issue #154 - Names of worksheets
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
      WHERE
          ( row_from <= ip_cell_row AND row_to >= ip_cell_row )
      AND
          ( col_from <= lv_column AND col_to >= lv_column ).

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
    DATA: lo_worksheets_iterator TYPE REF TO cl_object_collection_iterator,
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
          ls_sheet_content TYPE zexcel_s_cell_data.

    lv_column = zcl_excel_common=>convert_column2int( ip_column ).

    READ TABLE sheet_content INTO ls_sheet_content WITH TABLE KEY cell_row     = ip_row
                                                                  cell_column  = lv_column.

    ep_rc       = sy-subrc.
    ep_value    = ls_sheet_content-cell_value.
    ep_guid     = ls_sheet_content-cell_style.       " issue 139 - added this to be used for columnwidth calculation
    ep_formula  = ls_sheet_content-cell_formula.

    " Addition to solve issue #120, contribution by Stefan Schmöcker
    DATA: style_iterator TYPE REF TO cl_object_collection_iterator,
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
    eo_columns = me->columns.
  ENDMETHOD.                    "GET_COLUMNS


  METHOD get_columns_iterator.

    eo_iterator = me->columns->get_iterator( ).

  ENDMETHOD.                    "GET_COLUMNS_ITERATOR


  METHOD get_comments.
    DATA: lo_comment  TYPE REF TO zcl_excel_comment,
          lo_iterator TYPE REF TO cl_object_collection_iterator.

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
*    IF upper_cell-cell_coords = '0'.
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
          lo_iterator TYPE REF TO cl_object_collection_iterator.

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
    eo_rows = me->rows.
  ENDMETHOD.                    "GET_ROWS


  METHOD get_rows_iterator.

    eo_iterator = me->rows->get_iterator( ).

  ENDMETHOD.                    "GET_ROWS_ITERATOR


  METHOD get_row_outlines.

    rt_row_outlines = me->mt_row_outlines.

  ENDMETHOD.                    "GET_ROW_OUTLINES


  METHOD get_style_cond.

    DATA: lo_style_iterator TYPE REF TO cl_object_collection_iterator,
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
            REFRESH et_table.
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
*               LONG_TEXT  =
                output = l_value
*               SHORT_TEXT =
              EXCEPTIONS
                OTHERS = 1.
            IF sy-subrc <> 0.
* MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
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


  METHOD print_title_set_range.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmoecker,                            2012-12-02
*--------------------------------------------------------------------*


    DATA: lo_range_iterator         TYPE REF TO cl_object_collection_iterator,
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
          lv_row_end          TYPE zexcel_cell_row,
          lv_column_start     TYPE zexcel_cell_column_alpha,
          lv_column_end       TYPE zexcel_cell_column_alpha,
          lv_column_start_int TYPE zexcel_cell_column_alpha,
          lv_column_end_int   TYPE zexcel_cell_column_alpha.

    MOVE: ip_row_to TO lv_row_end,
          ip_row    TO lv_row.

    IF lv_row_end IS INITIAL OR ip_row_to IS NOT SUPPLIED.
      lv_row_end = lv_row.
    ENDIF.

    MOVE: ip_column_start TO lv_column_start,
          ip_column_end   TO lv_column_end.

    IF lv_column_end IS INITIAL OR ip_column_end IS NOT SUPPLIED.
      lv_column_end = lv_column_start.
    ENDIF.

    lv_column_start_int = zcl_excel_common=>convert_column2int( lv_column_start ).
    lv_column_end_int   = zcl_excel_common=>convert_column2int( lv_column_end ).

    IF lv_column_start_int > lv_column_end_int OR lv_row > lv_row_end.

      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'Wrong Merging Parameters'.

    ENDIF.

    IF ip_data_type IS SUPPLIED OR
       ip_abap_type IS SUPPLIED.

      me->set_cell( ip_column    = lv_column_start
                    ip_row       = lv_row
                    ip_value     = ip_value
                    ip_formula   = ip_formula
                    ip_style     = ip_style
                    ip_hyperlink = ip_hyperlink
                    ip_data_type = ip_data_type
                    ip_abap_type = ip_abap_type ).

    ELSE.

      me->set_cell( ip_column    = lv_column_start
                    ip_row       = lv_row
                    ip_value     = ip_value
                    ip_formula   = ip_formula
                    ip_style     = ip_style
                    ip_hyperlink = ip_hyperlink ).

    ENDIF.

    IF ip_style IS SUPPLIED.

      me->set_area_style( ip_column_start = lv_column_start
                          ip_column_end   = lv_column_end
                          ip_row          = lv_row
                          ip_row_to       = lv_row_end
                          ip_style        = ip_style ).
    ENDIF.

    IF ip_merge IS SUPPLIED AND ip_merge = abap_true.

      me->set_merge( ip_column_start = lv_column_start
                     ip_column_end   = lv_column_end
                     ip_row          = lv_row
                     ip_row_to       = lv_row_end ).

    ENDIF.

  ENDMETHOD.                    "set_area


  METHOD set_area_formula.
    DATA: ld_row            TYPE zexcel_cell_row,
          ld_row_end        TYPE zexcel_cell_row,
          ld_column         TYPE zexcel_cell_column_alpha,
          ld_column_end     TYPE zexcel_cell_column_alpha,
          ld_column_int     TYPE zexcel_cell_column_alpha,
          ld_column_end_int TYPE zexcel_cell_column_alpha.

    MOVE: ip_row_to TO ld_row_end,
          ip_row    TO ld_row.
    IF ld_row_end IS INITIAL OR ip_row_to IS NOT SUPPLIED.
      ld_row_end = ld_row.
    ENDIF.

    MOVE: ip_column_start TO ld_column,
          ip_column_end   TO ld_column_end.

    IF ld_column_end IS INITIAL OR ip_column_end IS NOT SUPPLIED.
      ld_column_end = ld_column.
    ENDIF.

    ld_column_int      = zcl_excel_common=>convert_column2int( ld_column ).
    ld_column_end_int  = zcl_excel_common=>convert_column2int( ld_column_end ).

    IF ld_column_int > ld_column_end_int OR ld_row > ld_row_end.
      RAISE EXCEPTION TYPE zcx_excel
        EXPORTING
          error = 'Wrong Merging Parameters'.
    ENDIF.

    me->set_cell_formula( ip_column = ld_column ip_row = ld_row
                          ip_formula = ip_formula ).

    IF ip_merge IS SUPPLIED AND ip_merge = abap_true.
      me->set_merge( ip_column_start = ld_column ip_row = ld_row
                     ip_column_end   = ld_column_end   ip_row_to = ld_row_end ).
    ENDIF.
  ENDMETHOD.                    "set_area_formula


  METHOD set_area_style.
    DATA: ld_row_start        TYPE zexcel_cell_row,
          ld_row_end          TYPE zexcel_cell_row,
          ld_column_start_int TYPE zexcel_cell_column,
          ld_column_end_int   TYPE zexcel_cell_column,
          ld_current_column   TYPE zexcel_cell_column_alpha,
          ld_current_row      TYPE zexcel_cell_row.

    MOVE: ip_row_to TO ld_row_end,
          ip_row    TO ld_row_start.
    IF ld_row_end IS INITIAL OR ip_row_to IS NOT SUPPLIED.
      ld_row_end = ld_row_start.
    ENDIF.
    ld_column_start_int = zcl_excel_common=>convert_column2int( ip_column_start ).
    ld_column_end_int   = zcl_excel_common=>convert_column2int( ip_column_end ).
    IF ld_column_end_int IS INITIAL OR ip_column_end IS NOT SUPPLIED.
      ld_column_end_int = ld_column_start_int.
    ENDIF.

    WHILE ld_column_start_int <= ld_column_end_int.
      ld_current_column = zcl_excel_common=>convert_column2alpha( ld_column_start_int ).
      ld_current_row = ld_row_start.
      WHILE ld_current_row <= ld_row_end.
        me->set_cell_style( ip_row = ld_current_row ip_column = ld_current_column
                            ip_style = ip_style ).
        ADD 1 TO ld_current_row.
      ENDWHILE.
      ADD 1 TO ld_column_start_int.
    ENDWHILE.
    IF ip_merge IS SUPPLIED AND ip_merge = abap_true.
      me->set_merge( ip_column_start = ip_column_start ip_row = ld_row_start
                     ip_column_end   = ld_current_column    ip_row_to = ld_row_end ).
    ENDIF.
  ENDMETHOD.                    "SET_AREA_STYLE


  METHOD set_area_hyperlink.
    DATA: ld_row_start        TYPE zexcel_cell_row,
          ld_row_end          TYPE zexcel_cell_row,
          ld_column_start_int TYPE zexcel_cell_column,
          ld_column_end_int   TYPE zexcel_cell_column,
          ld_current_column   TYPE zexcel_cell_column_alpha,
          ld_current_row      TYPE zexcel_cell_row,
          ld_value            TYPE string.
    DATA: lv_column    TYPE zexcel_cell_column,
          lo_hyperlink TYPE REF TO zcl_excel_hyperlink.

    MOVE: ip_row_to TO ld_row_end,
          ip_row    TO ld_row_start.
    IF ld_row_end IS INITIAL OR ip_row_to IS NOT SUPPLIED.
      ld_row_end = ld_row_start.
    ENDIF.
    ld_column_start_int = zcl_excel_common=>convert_column2int( ip_column_start ).
    ld_column_end_int   = zcl_excel_common=>convert_column2int( ip_column_end ).
    IF ld_column_end_int IS INITIAL OR ip_column_end IS NOT SUPPLIED.
      ld_column_end_int = ld_column_start_int.
    ENDIF.

    WHILE ld_column_start_int <= ld_column_end_int.
      ld_current_column = zcl_excel_common=>convert_column2alpha( ld_column_start_int ).
      ld_current_row = ld_row_start.
      WHILE ld_current_row <= ld_row_end.

        me->get_cell( EXPORTING ip_column = ld_current_column ip_row = ld_current_row
                      IMPORTING ep_value  = ld_value ).

        IF ip_is_internal = abap_true.
          lo_hyperlink = zcl_excel_hyperlink=>create_internal_link( iv_location = ip_url ).
        ELSE.
          lo_hyperlink = zcl_excel_hyperlink=>create_external_link( iv_url = ip_url ).
        ENDIF.

        me->set_cell( ip_column = ld_current_column ip_row = ld_current_row ip_value = ld_value ip_hyperlink = lo_hyperlink ).

        ADD 1 TO ld_current_row.
      ENDWHILE.
      ADD 1 TO ld_column_start_int.
    ENDWHILE.

  ENDMETHOD.                    "SET_AREA_HYPERLINK


  METHOD set_cell.

    DATA: lv_column        TYPE zexcel_cell_column,
          ls_sheet_content TYPE zexcel_s_cell_data,
          lv_row_alpha     TYPE string,
          lv_col_alpha     TYPE zexcel_cell_column_alpha,
          lv_value         TYPE zexcel_cell_value,
          lv_data_type     TYPE zexcel_cell_data_type,
          lv_value_type    TYPE abap_typekind,
          lv_style_guid    TYPE zexcel_cell_style,
          lo_addit         TYPE REF TO cl_abap_elemdescr,
          lo_value         TYPE REF TO data,
          lo_value_new     TYPE REF TO data.

    FIELD-SYMBOLS: <fs_sheet_content> TYPE zexcel_s_cell_data,
                   <fs_numeric>       TYPE numeric,
                   <fs_date>          TYPE d,
                   <fs_time>          TYPE t,
                   <fs_value>         TYPE simple,
                   <fs_typekind_int8> TYPE abap_typekind.
    FIELD-SYMBOLS: <fs_column_formula> TYPE mty_s_column_formula.


    IF ip_value  IS NOT SUPPLIED
        AND ip_formula IS NOT SUPPLIED
        AND ip_column_formula_id = 0.
      zcx_excel=>raise_text( 'Please provide the value or formula' ).
    ENDIF.

* Begin of change issue #152 - don't touch exisiting style if only value is passed
*  lv_style_guid = ip_style.
    lv_column = zcl_excel_common=>convert_column2int( ip_column ).
    IF ip_column_formula_id <> 0.
      check_cell_column_formula(
          it_column_formulas   = column_formulas
          ip_column_formula_id = ip_column_formula_id
          ip_formula           = ip_formula
          ip_value             = ip_value
          ip_row               = ip_row
          ip_column            = lv_column ).
    ENDIF.
    READ TABLE sheet_content ASSIGNING <fs_sheet_content> WITH TABLE KEY cell_row    = ip_row      " Changed to access via table key , Stefan Schmöcker, 2013-08-03
                                                                         cell_column = lv_column.
    IF sy-subrc = 0.
      IF ip_style IS INITIAL.
        " If no style is provided as method-parameter and cell is found use cell's current style
        lv_style_guid = <fs_sheet_content>-cell_style.
      ELSE.
        " Style provided as method-parameter --> use this
        lv_style_guid = ip_style.
      ENDIF.
    ELSE.
      " No cell found --> use supplied style even if empty
      lv_style_guid = ip_style.
    ENDIF.
* End of change issue #152 - don't touch exisiting style if only value is passed

    IF ip_value IS SUPPLIED.
      "if data type is passed just write the value. Otherwise map abap type to excel and perform conversion
      "IP_DATA_TYPE is passed by excel reader so source types are preserved
*First we get reference into local var.
      CREATE DATA lo_value LIKE ip_value.
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
            lo_addit = cl_abap_elemdescr=>get_i( ).
            CREATE DATA lo_value_new TYPE HANDLE lo_addit.
            ASSIGN lo_value_new->* TO <fs_numeric>.
            IF sy-subrc = 0.
              <fs_numeric> = <fs_value>.
              lv_value = zcl_excel_common=>number_to_excel_string( ip_value = <fs_numeric> ).
            ENDIF.

          WHEN cl_abap_typedescr=>typekind_float OR cl_abap_typedescr=>typekind_packed or
               cl_abap_typedescr=>typekind_decfloat OR
               cl_abap_typedescr=>typekind_decfloat16 OR
               cl_abap_typedescr=>typekind_decfloat34.
            lo_addit = cl_abap_elemdescr=>get_f( ).
            CREATE DATA lo_value_new TYPE HANDLE lo_addit.
            ASSIGN lo_value_new->* TO <fs_numeric>.
            IF sy-subrc = 0.
              <fs_numeric> = <fs_value>.
              lv_value = zcl_excel_common=>number_to_excel_string( ip_value = <fs_numeric> ).
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

    ENDIF.

    IF ip_hyperlink IS BOUND.
      ip_hyperlink->set_cell_reference( ip_column = ip_column
                                        ip_row = ip_row ).
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
*  lv_column = zcl_excel_common=>convert_column2int( ip_column ).

*  READ TABLE sheet_content ASSIGNING <fs_sheet_content> WITH KEY cell_row    = ip_row
*                                                                 cell_column = lv_column.
*
*  IF sy-subrc EQ 0.
    IF <fs_sheet_content> IS ASSIGNED.
* End of change issue #152 - don't touch exisiting style if only value is passed
      <fs_sheet_content>-cell_value   = lv_value.
      <fs_sheet_content>-cell_formula = ip_formula.
      <fs_sheet_content>-column_formula_id = ip_column_formula_id.
      <fs_sheet_content>-cell_style   = lv_style_guid.
      <fs_sheet_content>-data_type    = lv_data_type.
    ELSE.
      ls_sheet_content-cell_row     = ip_row.
      ls_sheet_content-cell_column  = lv_column.
      ls_sheet_content-cell_value   = lv_value.
      ls_sheet_content-cell_formula = ip_formula.
      ls_sheet_content-column_formula_id = ip_column_formula_id.
      ls_sheet_content-cell_style   = lv_style_guid.
      ls_sheet_content-data_type    = lv_data_type.
      lv_row_alpha = ip_row.
*    SHIFT lv_row_alpha RIGHT DELETING TRAILING space."del #152 - replaced with condense - should be faster
*    SHIFT lv_row_alpha LEFT DELETING LEADING space.  "del #152 - replaced with condense - should be faster
      CONDENSE lv_row_alpha NO-GAPS.                    "ins #152 - replaced 2 shifts      - should be faster
      lv_col_alpha = zcl_excel_common=>convert_column2alpha( ip_column ).       " issue #155 - less restrictive typing for ip_column
      CONCATENATE lv_col_alpha lv_row_alpha INTO ls_sheet_content-cell_coords.  " issue #155 - less restrictive typing for ip_column
      INSERT ls_sheet_content INTO TABLE sheet_content ASSIGNING <fs_sheet_content>. "ins #152 - Now <fs_sheet_content> always holds the data
*    APPEND ls_sheet_content TO sheet_content.
*    SORT sheet_content BY cell_row cell_column.
      " me->update_dimension_range( ).

    ENDIF.

* Begin of change issue #152 - don't touch exisiting style if only value is passed
* For Date- or Timefields change the formatcode if nothing is set yet
* Enhancement option:  Check if existing formatcode is a date/ or timeformat
*                      If not, use default
    DATA: lo_format_code_datetime TYPE zexcel_number_format.
    DATA: stylemapping    TYPE zexcel_s_stylemapping.
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
        me->change_cell_style( ip_column                      = ip_column
                               ip_row                         = ip_row
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
        me->change_cell_style( ip_column                      = ip_column
                               ip_row                         = ip_row
                               ip_number_format_format_code   = lo_format_code_datetime ).

    ENDCASE.
* End of change issue #152 - don't touch exisiting style if only value is passed

* Fix issue #162
    lv_value = ip_value.
    IF lv_value CS cl_abap_char_utilities=>cr_lf.
      me->change_cell_style( ip_column               = ip_column
                             ip_row                  = ip_row
                             ip_alignment_wraptext   = abap_true ).
    ENDIF.
* End of Fix issue #162

  ENDMETHOD.                    "SET_CELL


  METHOD set_cell_formula.
    DATA:
      lv_column        TYPE zexcel_cell_column,
      ls_sheet_content LIKE LINE OF me->sheet_content.

    FIELD-SYMBOLS:
                <sheet_content>                 LIKE LINE OF me->sheet_content.

*--------------------------------------------------------------------*
* Get cell to set formula into
*--------------------------------------------------------------------*
    lv_column = zcl_excel_common=>convert_column2int( ip_column ).
    READ TABLE me->sheet_content ASSIGNING <sheet_content> WITH TABLE KEY cell_row    = ip_row
                                                                          cell_column = lv_column.
    IF sy-subrc <> 0.                                                                           " Create new entry in sheet_content if necessary
      CHECK ip_formula IS INITIAL.                                                              " no need to create new entry in sheet_content when no formula is passed
      ls_sheet_content-cell_row    = ip_row.
      ls_sheet_content-cell_column = lv_column.
      INSERT ls_sheet_content INTO TABLE me->sheet_content ASSIGNING <sheet_content>.
    ENDIF.

*--------------------------------------------------------------------*
* Fieldsymbol now holds the relevant cell
*--------------------------------------------------------------------*
    <sheet_content>-cell_formula = ip_formula.


  ENDMETHOD.                    "SET_CELL_FORMULA


  METHOD set_cell_style.

    DATA: lv_column     TYPE zexcel_cell_column,
          lv_style_guid TYPE zexcel_cell_style.

    FIELD-SYMBOLS: <fs_sheet_content> TYPE zexcel_s_cell_data.

    lv_style_guid = ip_style.

    lv_column = zcl_excel_common=>convert_column2int( ip_column ).

    READ TABLE sheet_content ASSIGNING <fs_sheet_content> WITH KEY cell_row    = ip_row
                                                                   cell_column = lv_column.

    IF sy-subrc EQ 0.
      <fs_sheet_content>-cell_style   = lv_style_guid.
    ELSE.
      set_cell( ip_column = ip_column ip_row = ip_row ip_value = '' ip_style = ip_style ).
    ENDIF.

  ENDMETHOD.                    "SET_CELL_STYLE


  METHOD set_column_width.
    DATA: lo_column  TYPE REF TO zcl_excel_column.
    DATA: width             TYPE float.

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
          lv_errormessage TYPE string.

    ...
    "just after variables definition
    IF ip_value IS SUPPLIED OR ip_formula IS SUPPLIED.
      " if there is a value or formula set the value to the top-left cell
      "maybe it is necessary to support other paramters for set_cell
      IF ip_value IS SUPPLIED.
        me->set_cell( ip_row = ip_row ip_column = ip_column_start
                      ip_value = ip_value ).
      ENDIF.
      IF ip_formula IS SUPPLIED.
        me->set_cell( ip_row = ip_row ip_column = ip_column_start
                      ip_value = ip_formula ).
      ENDIF.
    ENDIF.
    "call to set_merge_style to apply the style to all cells at the matrix
    IF ip_style IS SUPPLIED.
      me->set_merge_style( ip_row = ip_row ip_column_start = ip_column_start
                           ip_row_to = ip_row_to ip_column_end = ip_column_end
                           ip_style = ip_style ).
    ENDIF.
    ...
*--------------------------------------------------------------------*
* Build new range area to insert into range table
*--------------------------------------------------------------------*
    ls_merge-row_from = ip_row.
    IF ip_row IS SUPPLIED AND ip_row IS NOT INITIAL AND ip_row_to IS NOT SUPPLIED.
      ls_merge-row_to   = ls_merge-row_from.
    ELSE.
      ls_merge-row_to   = ip_row_to.
    ENDIF.
    IF ls_merge-row_from > ls_merge-row_to.
      lv_errormessage = 'Merge: First row larger then last row'(405).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

    ls_merge-col_from = zcl_excel_common=>convert_column2int( ip_column_start ).
    IF ip_column_start IS SUPPLIED AND ip_column_start IS NOT INITIAL AND ip_column_end IS NOT SUPPLIED.
      ls_merge-col_to   = ls_merge-col_from.
    ELSE.
      ls_merge-col_to   = zcl_excel_common=>convert_column2int( ip_column_end ).
    ENDIF.
    IF ls_merge-col_from > ls_merge-col_to.
      lv_errormessage = 'Merge: First column larger then last column'(406).
      zcx_excel=>raise_text( lv_errormessage ).
    ENDIF.

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
          ld_column_start   TYPE zexcel_cell_column,
          ld_column_end     TYPE zexcel_cell_column,
          ld_current_column TYPE zexcel_cell_column_alpha,
          ld_current_row    TYPE zexcel_cell_row.

    MOVE: ip_row_to TO ld_row_end,
          ip_row    TO ld_row_start.
    IF ld_row_end IS INITIAL.
      ld_row_end = ld_row_start.
    ENDIF.
    ld_column_start = zcl_excel_common=>convert_column2int( ip_column_start ).
    ld_column_end   = zcl_excel_common=>convert_column2int( ip_column_end ).
    IF ld_column_end IS INITIAL.
      ld_column_end = ld_column_start.
    ENDIF.
    "set the style cell by cell
    WHILE ld_column_start <= ld_column_end.
      ld_current_column = zcl_excel_common=>convert_column2alpha( ld_column_start ).
      ld_current_row = ld_row_start.
      WHILE ld_current_row <= ld_row_end.
        me->set_cell_style( ip_row = ld_current_row ip_column = ld_current_column
                            ip_style = ip_style ).
        ADD 1 TO ld_current_row.
      ENDWHILE.
      ADD 1 TO ld_column_start.
    ENDWHILE.
  ENDMETHOD.                    "set_merge_style


  METHOD set_print_gridlines.
    me->print_gridlines = i_print_gridlines.
  ENDMETHOD.                    "SET_PRINT_GRIDLINES


  METHOD set_row_height.
    DATA: lo_row  TYPE REF TO zcl_excel_row.
    DATA: height  TYPE float.

    lo_row = me->get_row( ip_row ).

* if a fix size is supplied use this
    TRY.
        height = ip_height_fix.
        IF height <= 0.
          zcx_excel=>raise_text( 'Please supply a positive number as row-height' ).
        ENDIF.
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

    DATA: lo_tabdescr     TYPE REF TO cl_abap_structdescr,
          lr_data         TYPE REF TO data,
          ls_header       TYPE x030l,
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

    lo_tabdescr ?= cl_abap_structdescr=>describe_by_data_ref( lr_data ).

    ls_header = lo_tabdescr->get_ddic_header( ).

    lt_dfies = lo_tabdescr->get_ddic_field_list( ).

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
        MOVE <fs_fldval> TO lv_cell_value.
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
    DATA: lo_worksheets_iterator TYPE REF TO cl_object_collection_iterator,
          lo_worksheet           TYPE REF TO zcl_excel_worksheet,
          errormessage           TYPE string,
          lv_rangesheetname_old  TYPE string,
          lv_rangesheetname_new  TYPE string,
          lo_ranges_iterator     TYPE REF TO cl_object_collection_iterator,
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

    lv_row_alpha = upper_cell-cell_row.
    lv_column_alpha = zcl_excel_common=>convert_column2alpha( upper_cell-cell_column ).
    SHIFT lv_row_alpha RIGHT DELETING TRAILING space.
    SHIFT lv_row_alpha LEFT DELETING LEADING space.
    CONCATENATE lv_column_alpha lv_row_alpha INTO upper_cell-cell_coords.

    lv_row_alpha = lower_cell-cell_row.
    lv_column_alpha = zcl_excel_common=>convert_column2alpha( lower_cell-cell_column ).
    SHIFT lv_row_alpha RIGHT DELETING TRAILING space.
    SHIFT lv_row_alpha LEFT DELETING LEADING space.
    CONCATENATE lv_column_alpha lv_row_alpha INTO lower_cell-cell_coords.

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
