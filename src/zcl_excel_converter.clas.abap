class ZCL_EXCEL_CONVERTER definition
  public
  create public .

*"* public components of class ZCL_EXCEL_CONVERTER
*"* do not include other source files here!!!
public section.
  type-pools ABAP .

  class-methods CLASS_CONSTRUCTOR .
  methods ASK_OPTION
    returning
      value(RS_OPTION) type ZEXCEL_S_CONVERTER_OPTION
    raising
      ZCX_EXCEL .
  methods CONVERT
    importing
      !IS_OPTION type ZEXCEL_S_CONVERTER_OPTION optional
      !IO_ALV type ref to OBJECT optional
      !IT_TABLE type STANDARD TABLE optional
      !I_ROW_INT type I default 1
      !I_COLUMN_INT type I default 1
      !I_TABLE type FLAG optional
      !I_STYLE_TABLE type ZEXCEL_TABLE_STYLE optional
      !IO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET optional
    changing
      !CO_EXCEL type ref to ZCL_EXCEL optional
    raising
      ZCX_EXCEL .
  methods CREATE_PATH
    returning
      value(R_PATH) type STRING .
  methods GET_FILE
    exporting
      !E_BYTECOUNT type I
      !ET_FILE type SOLIX_TAB
      !E_FILE type XSTRING .
  methods GET_OPTION
    returning
      value(RS_OPTION) type ZEXCEL_S_CONVERTER_OPTION .
  methods OPEN_FILE .
  methods SET_OPTION
    importing
      !IS_OPTION type ZEXCEL_S_CONVERTER_OPTION .
  methods WRITE_FILE
    importing
      !I_PATH type STRING optional .
*"* protected components of class ZCL_EXCEL_CONVERTER
*"* do not include other source files here!!!
protected section.

  types:
    BEGIN OF t_relationship,
           id     TYPE string,
           type   TYPE string,
           target TYPE string,
         END OF t_relationship .
  types:
    BEGIN OF t_fileversion,
            appname       TYPE string,
            lastedited    TYPE string,
            lowestedited  TYPE string,
            rupbuild      TYPE string,
            codename      TYPE string,
       END OF t_fileversion .
  types:
    BEGIN OF t_sheet,
              name    TYPE string,
              sheetid TYPE string,
              id      TYPE string,
         END OF t_sheet .
  types:
    BEGIN OF t_workbookpr,
              codename            TYPE string,
              defaultthemeversion TYPE string,
         END OF t_workbookpr .
  types:
    BEGIN OF t_sheetpr,
              codename            TYPE string,
         END OF t_sheetpr .

  data W_ROW_INT type ZEXCEL_CELL_ROW value 1. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
  data W_COL_INT type ZEXCEL_CELL_COLUMN value 1. "#EC NOTEXT .  .  .  .  .  .  .  .  .  . " .
*"* private components of class ZCL_EXCEL_CONVERTER
*"* do not include other source files here!!!
private section.
  CLASS-METHODS get_subclasses
    IMPORTING
      is_clskey  TYPE seoclskey
    CHANGING
      ct_classes TYPE seor_implementing_keys.

  data WO_EXCEL type ref to ZCL_EXCEL .
  data WO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET .
  data WO_AUTOFILTER type ref to ZCL_EXCEL_AUTOFILTER .
  data WO_TABLE type ref to DATA .
  data WO_DATA type ref to DATA .
  data WT_FIELDCATALOG type ZEXCEL_T_CONVERTER_FCAT .
  data WS_LAYOUT type ZEXCEL_S_CONVERTER_LAYO .
  data WT_COLORS type ZEXCEL_T_CONVERTER_COL .
  data WT_FILTER type ZEXCEL_T_CONVERTER_FIL .
  class-data WT_OBJECTS type TT_ALV_TYPES .
  class-data W_FCOUNT type NUMC3 .
  data WT_SORT_VALUES type TT_SORT_VALUES .
  data WT_SUBTOTAL_ROWS type TT_SUBTOTAL_ROWS .
  data WT_STYLES type TT_STYLES .
  constants C_TYPE_HDR type CHAR1 value 'H'. "#EC NOTEXT
  constants C_TYPE_STR type CHAR1 value 'P'. "#EC NOTEXT
  constants C_TYPE_NOR type CHAR1 value 'N'. "#EC NOTEXT
  constants C_TYPE_SUB type CHAR1 value 'S'. "#EC NOTEXT
  constants C_TYPE_TOT type CHAR1 value 'T'. "#EC NOTEXT
  data WT_COLOR_STYLES type TT_COLOR_STYLES .
  class-data WS_OPTION type ZEXCEL_S_CONVERTER_OPTION .
  class-data WS_INDX type INDX .

  class-methods INIT_OPTION .
  class-methods GET_ALV_CONVERTERS.
  methods BIND_TABLE
    importing
      !I_STYLE_TABLE type ZEXCEL_TABLE_STYLE
    returning
      value(R_FREEZE_COL) type INT1
    raising
      ZCX_EXCEL .
  methods BIND_CELLS
    returning
      value(R_FREEZE_COL) type INT1
    raising
      ZCX_EXCEL .
  methods CLEAN_FIELDCATALOG .
  methods CREATE_COLOR_STYLE
    importing
      !I_STYLE type ZEXCEL_CELL_STYLE
      !IS_COLORS type ZEXCEL_S_CONVERTER_COL
    returning
      value(RO_STYLE) type ref to ZCL_EXCEL_STYLE .
  methods CREATE_FORMULAR_SUBTOTAL
    importing
      !I_ROW_INT_START type ZEXCEL_CELL_ROW
      !I_ROW_INT_END type ZEXCEL_CELL_ROW
      !I_COLUMN type ZEXCEL_CELL_COLUMN_ALPHA
      !I_TOTALS_FUNCTION type ZEXCEL_TABLE_TOTALS_FUNCTION
    returning
      value(R_FORMULA) type STRING .
  methods CREATE_FORMULAR_TOTAL
    importing
      !I_ROW_INT type ZEXCEL_CELL_ROW
      !I_COLUMN type ZEXCEL_CELL_COLUMN_ALPHA
      !I_TOTALS_FUNCTION type ZEXCEL_TABLE_TOTALS_FUNCTION
    returning
      value(R_FORMULA) type STRING .
  methods CREATE_STYLE_HDR
    importing
      !I_ALIGNMENT type ZEXCEL_ALIGNMENT optional
    returning
      value(RO_STYLE) type ref to ZCL_EXCEL_STYLE .
  methods CREATE_STYLE_NORMAL
    importing
      !I_ALIGNMENT type ZEXCEL_ALIGNMENT optional
      !I_INTTYPE type INTTYPE optional
      !I_DECIMALS type INT1 optional
    returning
      value(RO_STYLE) type ref to ZCL_EXCEL_STYLE .
  methods CREATE_STYLE_STRIPPED
    importing
      !I_ALIGNMENT type ZEXCEL_ALIGNMENT optional
      !I_INTTYPE type INTTYPE optional
      !I_DECIMALS type INT1 optional
    returning
      value(RO_STYLE) type ref to ZCL_EXCEL_STYLE .
  methods CREATE_STYLE_SUBTOTAL
    importing
      !I_ALIGNMENT type ZEXCEL_ALIGNMENT optional
      !I_INTTYPE type INTTYPE optional
      !I_DECIMALS type INT1 optional
    returning
      value(RO_STYLE) type ref to ZCL_EXCEL_STYLE .
  methods CREATE_STYLE_TOTAL
    importing
      !I_ALIGNMENT type ZEXCEL_ALIGNMENT optional
      !I_INTTYPE type INTTYPE optional
      !I_DECIMALS type INT1 optional
    returning
      value(RO_STYLE) type ref to ZCL_EXCEL_STYLE .
  methods CREATE_TABLE .
  methods CREATE_TEXT_SUBTOTAL
    importing
      !I_VALUE type ANY
      !I_TOTALS_FUNCTION type ZEXCEL_TABLE_TOTALS_FUNCTION
    returning
      value(R_TEXT) type STRING .
  methods CREATE_WORKSHEET
    importing
      !I_TABLE type FLAG default 'X'
      !I_STYLE_TABLE type ZEXCEL_TABLE_STYLE
    raising
      ZCX_EXCEL .
  methods EXECUTE_CONVERTER
    importing
      !IO_OBJECT type ref to OBJECT
      !IT_TABLE type STANDARD TABLE
    raising
      ZCX_EXCEL .
  methods GET_COLOR_STYLE
    importing
      !I_ROW type ZEXCEL_CELL_ROW
      !I_FIELDNAME type FIELDNAME
      !I_STYLE type ZEXCEL_CELL_STYLE
    returning
      value(R_STYLE) type ZEXCEL_CELL_STYLE .
  methods GET_FUNCTION_NUMBER
    importing
      !I_TOTALS_FUNCTION type ZEXCEL_TABLE_TOTALS_FUNCTION
    returning
      value(R_FUNCTION_NUMBER) type INT1 .
  methods GET_STYLE
    importing
      !I_TYPE type CHAR1
      !I_ALIGNMENT type ZEXCEL_ALIGNMENT default SPACE
      !I_INTTYPE type INTTYPE default SPACE
      !I_DECIMALS type INT1 default 0
    returning
      value(R_STYLE) type ZEXCEL_CELL_STYLE .
  methods LOOP_NORMAL
    importing
      !I_ROW_INT type ZEXCEL_CELL_ROW
      !I_COL_INT type ZEXCEL_CELL_COLUMN
    returning
      value(R_FREEZE_COL) type INT1
    raising
      ZCX_EXCEL .
  methods LOOP_SUBTOTAL
    importing
      !I_ROW_INT type ZEXCEL_CELL_ROW
      !I_COL_INT type ZEXCEL_CELL_COLUMN
    returning
      value(R_FREEZE_COL) type INT1
    raising
      ZCX_EXCEL .
  methods SET_AUTOFILTER_AREA .
  methods SET_CELL_FORMAT
    importing
      !I_INTTYPE type INTTYPE
      !I_DECIMALS type INT1
    returning
      value(R_FORMAT) type ZEXCEL_NUMBER_FORMAT .
  methods SET_FIELDCATALOG .
ENDCLASS.



CLASS ZCL_EXCEL_CONVERTER IMPLEMENTATION.


METHOD ask_option.
  DATA: ls_sval TYPE sval,
        lt_sval TYPE STANDARD TABLE OF sval,
        l_returncode TYPE string,
        lt_fields    TYPE ddfields,
        ls_fields    TYPE dfies.

  FIELD-SYMBOLS: <fs> TYPE any.

  rs_option = ws_option.

  CALL FUNCTION 'DDIF_FIELDINFO_GET'
    EXPORTING
      tabname        = 'ZEXCEL_S_CONVERTER_OPTION'
    TABLES
      dfies_tab      = lt_fields
    EXCEPTIONS
      not_found      = 1
      internal_error = 2
      OTHERS         = 3.
  IF sy-subrc <> 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
            WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ENDIF.

  LOOP AT lt_fields INTO ls_fields.
    ASSIGN COMPONENT ls_fields-fieldname OF STRUCTURE ws_option TO <fs>.
    IF sy-subrc = 0.
      CLEAR ls_sval.
      ls_sval-tabname      = ls_fields-tabname.
      ls_sval-fieldname    = ls_fields-fieldname.
      ls_sval-value        = <fs>.
      ls_sval-field_attr   = space.
      ls_sval-field_obl    = space.
      ls_sval-comp_code    = space.
      ls_sval-fieldtext    = ls_fields-scrtext_m.
      ls_sval-comp_tab     = space.
      ls_sval-comp_field   = space.
      ls_sval-novaluehlp   = space.
      INSERT ls_sval INTO TABLE lt_sval.
    ENDIF.
  ENDLOOP.

  CALL FUNCTION 'POPUP_GET_VALUES'
    EXPORTING
      popup_title     = 'Excel creation options'(008)
    IMPORTING
      returncode      = l_returncode
    TABLES
      fields          = lt_sval
    EXCEPTIONS
      error_in_fields = 1
      OTHERS          = 2.
  IF sy-subrc <> 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
            WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ELSEIF l_returncode = 'A'.
      RAISE EXCEPTION TYPE zcx_excel.
  ELSE.
    LOOP AT lt_sval INTO ls_sval.
      ASSIGN COMPONENT ls_sval-fieldname OF STRUCTURE ws_option TO <fs>.
      IF sy-subrc = 0.
        <fs> = ls_sval-value.
      ENDIF.
    ENDLOOP.
    set_option( is_option = ws_option ) .
    rs_option = ws_option.
  ENDIF.
ENDMETHOD.


method BIND_CELLS.

* Do we need subtotals with grouping
  READ TABLE wt_fieldcatalog TRANSPORTING NO FIELDS WITH KEY is_subtotalled = abap_true.
  IF sy-subrc = 0  .
    r_freeze_col = loop_subtotal( i_row_int = w_row_int
                                  i_col_int  = w_col_int ) .
  ELSE.
    r_freeze_col = loop_normal( i_row_int = w_row_int
                                i_col_int = w_col_int ) .
  ENDIF.

  endmethod.


METHOD bind_table.
  DATA: lt_field_catalog  TYPE zexcel_t_fieldcatalog,
        ls_field_catalog  TYPE zexcel_s_fieldcatalog,
        ls_fcat           TYPE zexcel_s_converter_fcat,
        lo_column         TYPE REF TO zcl_excel_column,
        lv_col_int        TYPE zexcel_cell_column,
        lv_col_alpha      TYPE zexcel_cell_column_alpha,
        ls_settings       TYPE zexcel_s_table_settings,
        lv_line           TYPE i.

  FIELD-SYMBOLS: <fs_tab>         TYPE ANY TABLE.

  ASSIGN wo_data->* TO <fs_tab> .

  ls_settings-table_style      = i_style_table.
  ls_settings-top_left_column  = zcl_excel_common=>convert_column2alpha( ip_column = w_col_int ).
  ls_settings-top_left_row     = w_row_int.
  ls_settings-show_row_stripes = ws_layout-is_stripped.

  DESCRIBE TABLE  wt_fieldcatalog  LINES lv_line.
  lv_line = lv_line + 1 + w_col_int.
  ls_settings-bottom_right_column = zcl_excel_common=>convert_column2alpha( ip_column = lv_line ).

  DESCRIBE TABLE <fs_tab> LINES lv_line.
  ls_settings-bottom_right_row = lv_line + 1 + w_row_int.
  SORT wt_fieldcatalog BY position.
  LOOP AT wt_fieldcatalog INTO ls_fcat.
    MOVE-CORRESPONDING ls_fcat TO ls_field_catalog.
    ls_field_catalog-dynpfld = abap_true.
    INSERT ls_field_catalog INTO TABLE lt_field_catalog.
  ENDLOOP.

  wo_worksheet->bind_table(
    EXPORTING
      ip_table          = <fs_tab>
      it_field_catalog  = lt_field_catalog
      is_table_settings = ls_settings
    IMPORTING
      es_table_settings = ls_settings
         ).
  LOOP AT wt_fieldcatalog INTO ls_fcat.
    lv_col_int = w_col_int + ls_fcat-position - 1.
    lv_col_alpha = zcl_excel_common=>convert_column2alpha( lv_col_int ).
* Freeze panes
    IF ls_fcat-fix_column = abap_true.
      ADD 1 TO r_freeze_col.
    ENDIF.
* Now let's check for optimized
    IF ls_fcat-is_optimized = abap_true.
      lo_column = wo_worksheet->get_column( ip_column = lv_col_alpha ).
      lo_column->set_auto_size( ip_auto_size = abap_true ) .
    ENDIF.
* Now let's check for visible
    IF ls_fcat-is_hidden = abap_true.
      lo_column = wo_worksheet->get_column( ip_column = lv_col_alpha ).
      lo_column->set_visible( ip_visible = abap_false ) .
    ENDIF.
  ENDLOOP.

ENDMETHOD.


method CLASS_CONSTRUCTOR.
  DATA: ls_objects TYPE ts_alv_types.
  DATA: ls_option TYPE zexcel_s_converter_option,
        l_uname TYPE sy-uname.

  GET PARAMETER ID 'ZUS' FIELD l_uname.
  IF l_uname IS INITIAL OR l_uname = space.
    l_uname = sy-uname.
  ENDIF.

  get_alv_converters( ).

  CONCATENATE 'EXCEL_' sy-uname INTO ws_indx-srtfd.

  IMPORT p1 = ls_option FROM DATABASE indx(xl) TO ws_indx ID ws_indx-srtfd.

  IF sy-subrc = 0.
    ws_option = ls_option.
  ELSE.
    init_option( ) .
  ENDIF.

  endmethod.


method CLEAN_FIELDCATALOG.
  DATA: l_position TYPE tabfdpos.

  FIELD-SYMBOLS: <fs_sfcat>   TYPE zexcel_s_converter_fcat.

  SORT wt_fieldcatalog BY position col_id.

  CLEAR l_position.
  LOOP AT wt_fieldcatalog ASSIGNING <fs_sfcat>.
    ADD 1 TO l_position.
    <fs_sfcat>-position = l_position.
* Default stype with alignment and format
    <fs_sfcat>-style_hdr      = get_style( i_type      = c_type_hdr
                                           i_alignment = <fs_sfcat>-alignment ).
    IF ws_layout-is_stripped = abap_true.
      <fs_sfcat>-style_stripped = get_style( i_type      = c_type_str
                                             i_alignment = <fs_sfcat>-alignment
                                             i_inttype   = <fs_sfcat>-inttype
                                             i_decimals  = <fs_sfcat>-decimals   ).
    ENDIF.
    <fs_sfcat>-style_normal   = get_style( i_type      = c_type_nor
                                           i_alignment = <fs_sfcat>-alignment
                                           i_inttype   = <fs_sfcat>-inttype
                                           i_decimals  = <fs_sfcat>-decimals   ).
    <fs_sfcat>-style_subtotal = get_style( i_type      = c_type_sub
                                           i_alignment = <fs_sfcat>-alignment
                                           i_inttype   = <fs_sfcat>-inttype
                                           i_decimals  = <fs_sfcat>-decimals   ).
    <fs_sfcat>-style_total    = get_style( i_type      = c_type_tot
                                           i_alignment = <fs_sfcat>-alignment
                                           i_inttype   = <fs_sfcat>-inttype
                                           i_decimals  = <fs_sfcat>-decimals   ).
  ENDLOOP.

  endmethod.


method CONVERT.

  IF is_option IS SUPPLIED.
    ws_option = is_option.
  ENDIF.

  execute_converter( EXPORTING io_object   = io_alv
                               it_table    = it_table ) .

  IF io_worksheet IS SUPPLIED AND io_worksheet IS BOUND.
    wo_worksheet = io_worksheet.
  ENDIF.
  IF co_excel IS SUPPLIED.
    IF co_excel IS NOT BOUND.
      CREATE OBJECT co_excel.
      co_excel->zif_excel_book_properties~creator = sy-uname.
    ENDIF.
    wo_excel = co_excel.
  ENDIF.

* Move table to data object and clean it up
  IF wt_fieldcatalog IS NOT INITIAL.
    create_table( ).
  ELSE.
    wo_data = wo_table .
  ENDIF.

  IF wo_excel IS NOT BOUND.
    CREATE OBJECT wo_excel.
    wo_excel->zif_excel_book_properties~creator = sy-uname.
  ENDIF.
  IF wo_worksheet IS NOT BOUND.
    " Get active sheet
    wo_worksheet = wo_excel->get_active_worksheet( ).
    wo_worksheet->set_title( ip_title = 'Sheet1'(001) ).
  ENDIF.

  IF i_row_int <= 0.
    w_row_int = 1.
  ELSE.
    w_row_int = i_row_int.
  ENDIF.
  IF i_column_int <= 0.
    w_col_int = 1.
  ELSE.
    w_col_int = i_column_int.
  ENDIF.

  create_worksheet( i_table       = i_table
                    i_style_table = i_style_table ) .

  endmethod.


method CREATE_COLOR_STYLE.
  DATA: ls_styles TYPE ts_styles.
  DATA: lo_style TYPE REF TO zcl_excel_style.

  READ TABLE wt_styles INTO ls_styles WITH KEY guid = i_style.
  IF sy-subrc = 0.
    lo_style                 = wo_excel->add_new_style( ).
*    lo_style->borders        = ls_styles-style->borders.
*    lo_style->protection     = ls_styles-style->protection.
    lo_style->font->bold                 = ls_styles-style->font->bold.
    lo_style->alignment->horizontal      = ls_styles-style->alignment->horizontal.
    lo_style->number_format->format_code = ls_styles-style->number_format->format_code.

    lo_style->font->color-rgb      = is_colors-fontcolor.
    lo_style->fill->filltype       = zcl_excel_style_fill=>c_fill_solid.
    lo_style->fill->fgcolor-rgb    = is_colors-fillcolor.

    ro_style = lo_style.
  ENDIF.
  endmethod.


method CREATE_FORMULAR_SUBTOTAL.
  data: l_row_alpha_start type string,
        l_row_alpha_end   type string,
        l_func_num        type string.

  l_row_alpha_start   = i_row_int_start.
  l_row_alpha_end     = i_row_int_end.

  l_func_num = get_function_number( i_totals_function = i_totals_function ).
  concatenate 'SUBTOTAL(' l_func_num ',' i_column l_row_alpha_start ':' i_column l_row_alpha_end ')' into r_formula.
  endmethod.


method CREATE_FORMULAR_TOTAL.
  data: l_row_alpha    type string,
        l_row_e_alpha  type string.

  l_row_alpha   = w_row_int + 1.
  l_row_e_alpha = i_row_int.

  concatenate i_totals_function '(' i_column l_row_alpha ':' i_column l_row_e_alpha ')' into r_formula.
  endmethod.


method CREATE_PATH.
  DATA:   l_sep                   TYPE c ,
          l_path                  TYPE string,
          l_return                TYPE i .

  CLEAR r_path.

  " Save the file
  cl_gui_frontend_services=>get_sapgui_workdir(
    CHANGING
      sapworkdir            = l_path
        EXCEPTIONS
          get_sapworkdir_failed = 1
          cntl_error            = 2
          error_no_gui          = 3
          not_supported_by_gui  = 4
         ).
  IF sy-subrc <> 0.
*       MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*                  WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
    CONCATENATE 'Excel_' w_fcount '.xlsx' INTO r_path.
  ELSE.
    DO.
      ADD 1 TO w_fcount.
*-obtain file separator character---------------------------------------
      CALL METHOD cl_gui_frontend_services=>get_file_separator
        CHANGING
          file_separator       = l_sep
        EXCEPTIONS
          cntl_error           = 1
          error_no_gui         = 2
          not_supported_by_gui = 3
          OTHERS               = 4.

      IF sy-subrc <> 0.
        l_sep = ''.
      ENDIF.

      CONCATENATE l_path l_sep 'Excel_' w_fcount '.xlsx' INTO r_path.

      IF cl_gui_frontend_services=>file_exist( file  = r_path ) = abap_true.
        cl_gui_frontend_services=>file_delete( EXPORTING filename = r_path
                                               CHANGING  rc       = l_return
                                               EXCEPTIONS OTHERS  = 1 ).
        IF sy-subrc = 0 .
          RETURN.
        ENDIF.
      ELSE.
        RETURN.
      ENDIF.
    ENDDO.
  ENDIF.

  endmethod.


method CREATE_STYLE_HDR.
  data: lo_style type ref to zcl_excel_style.

  lo_style                       = wo_excel->add_new_style( ).
  lo_style->font->bold           = abap_true.
  lo_style->font->color-rgb      = zcl_excel_style_color=>c_white.
  lo_style->fill->filltype       = zcl_excel_style_fill=>c_fill_solid.
  lo_style->fill->fgcolor-rgb    = 'FF4F81BD'.
  if i_alignment is supplied and i_alignment is not initial.
    lo_style->alignment->horizontal = i_alignment.
  endif.
  ro_style = lo_style .
  endmethod.


method CREATE_STYLE_NORMAL.
  DATA:   lo_style    TYPE REF TO zcl_excel_style,
          l_format    TYPE zexcel_number_format.

  IF i_inttype IS SUPPLIED AND i_inttype IS NOT INITIAL.
    l_format = set_cell_format(  i_inttype  = i_inttype
                                 i_decimals = i_decimals ) .
  ENDIF.
  IF l_format IS NOT INITIAL OR
     ( i_alignment IS SUPPLIED AND i_alignment IS NOT INITIAL ) .

    lo_style                       = wo_excel->add_new_style( ).

    IF i_alignment IS SUPPLIED AND i_alignment IS NOT INITIAL.
      lo_style->alignment->horizontal = i_alignment.
    ENDIF.

    IF l_format IS NOT INITIAL.
      lo_style->number_format->format_code = l_format.
    ENDIF.

    ro_style = lo_style .

  ENDIF.
  endmethod.


method CREATE_STYLE_STRIPPED.
  data:   lo_style    type ref to zcl_excel_style.
  data:  l_format    type zexcel_number_format.

  lo_style                       = wo_excel->add_new_style( ).
  lo_style->fill->filltype       = zcl_excel_style_fill=>c_fill_solid.
  lo_style->fill->fgcolor-rgb    = 'FFDBE5F1'.
  if i_alignment is supplied and i_alignment is not initial.
    lo_style->alignment->horizontal = i_alignment.
  endif.
  if i_inttype is supplied and i_inttype is not initial.
    l_format = set_cell_format(  i_inttype  = i_inttype
                                 i_decimals = i_decimals ) .
    if l_format is not initial.
      lo_style->number_format->format_code = l_format.
    endif.
  endif.
  ro_style = lo_style.

  endmethod.


method CREATE_STYLE_SUBTOTAL.
  data:   lo_style    type ref to zcl_excel_style.
  data:  l_format    type zexcel_number_format.

  lo_style                       = wo_excel->add_new_style( ).
  lo_style->font->bold           = abap_true.

  if i_alignment is supplied and i_alignment is not initial.
    lo_style->alignment->horizontal = i_alignment.
  endif.
  if i_inttype is supplied and i_inttype is not initial.
    l_format = set_cell_format(  i_inttype  = i_inttype
                                 i_decimals = i_decimals ) .
    if l_format is not initial.
      lo_style->number_format->format_code = l_format.
    endif.
  endif.

  ro_style = lo_style .

  endmethod.


method CREATE_STYLE_TOTAL.
  DATA:   lo_style    TYPE REF TO zcl_excel_style.
  DATA:  l_format    TYPE zexcel_number_format.

  lo_style                                   = wo_excel->add_new_style( ).
  lo_style->font->bold                       = abap_true.

  CREATE OBJECT lo_style->borders->top.
  lo_style->borders->top->border_style       = zcl_excel_style_border=>c_border_thin.
  lo_style->borders->top->border_color-rgb   = zcl_excel_style_color=>c_black.

  CREATE OBJECT lo_style->borders->right.
  lo_style->borders->right->border_style       = zcl_excel_style_border=>c_border_none.
  lo_style->borders->right->border_color-rgb   = zcl_excel_style_color=>c_black.

  CREATE OBJECT lo_style->borders->down.
  lo_style->borders->down->border_style      = zcl_excel_style_border=>c_border_double.
  lo_style->borders->down->border_color-rgb  = zcl_excel_style_color=>c_black.

  CREATE OBJECT lo_style->borders->left.
  lo_style->borders->left->border_style       = zcl_excel_style_border=>c_border_none.
  lo_style->borders->left->border_color-rgb   = zcl_excel_style_color=>c_black.

  IF i_alignment IS SUPPLIED AND i_alignment IS NOT INITIAL.
    lo_style->alignment->horizontal = i_alignment.
  ENDIF.
  IF i_inttype IS SUPPLIED AND i_inttype IS NOT INITIAL.
    l_format = set_cell_format(  i_inttype  = i_inttype
                                 i_decimals = i_decimals ) .
    IF l_format IS NOT INITIAL.
      lo_style->number_format->format_code = l_format.
    ENDIF.
  ENDIF.

  ro_style = lo_style .

  endmethod.


method CREATE_TABLE.
  TYPES: BEGIN OF ts_output,
           fieldname TYPE fieldname,
           function  TYPE funcname,
         END OF ts_output.

  DATA: lo_data TYPE REF TO data.
  DATA: lo_addit                TYPE REF TO cl_abap_elemdescr,
        lt_components_tab       TYPE cl_abap_structdescr=>component_table,
        ls_components           TYPE abap_componentdescr,
        lo_table                TYPE REF TO cl_abap_tabledescr,
        lo_struc                TYPE REF TO cl_abap_structdescr.

  FIELD-SYMBOLS: <fs_scat>  TYPE zexcel_s_converter_fcat,
                 <fs_stab>  TYPE ANY,
                 <fs_ttab>  TYPE STANDARD TABLE,
                 <fs>       TYPE ANY,
                 <fs_table> TYPE STANDARD TABLE.

  SORT wt_fieldcatalog BY position.
  ASSIGN wo_table->* TO <fs_table>.

  READ TABLE <fs_table> ASSIGNING <fs_stab> INDEX 1.
  IF sy-subrc EQ 0 .
    LOOP AT wt_fieldcatalog ASSIGNING <fs_scat>.
      ASSIGN COMPONENT <fs_scat>-columnname OF STRUCTURE <fs_stab> TO <fs>.
      IF sy-subrc = 0.
        ls_components-name   = <fs_scat>-columnname.
        TRY.
            lo_addit            ?= cl_abap_typedescr=>describe_by_data( <fs> ).
          CATCH cx_sy_move_cast_error.
            CLEAR lo_addit.
            DELETE TABLE wt_fieldcatalog FROM <fs_scat>.
        ENDTRY.
        IF lo_addit IS BOUND.
          ls_components-type   = lo_addit           .
          INSERT ls_components INTO TABLE lt_components_tab.
        ENDIF.
      ENDIF.
    ENDLOOP.
    IF lt_components_tab IS NOT INITIAL.
      "create new line type
      TRY.
          lo_struc = cl_abap_structdescr=>create( P_COMPONENTS = lt_components_tab
                                                  P_STRICT     = abap_false ).
        CATCH cx_sy_struct_creation.
          RETURN.  " We can not do anything in this case.
      ENDTRY.

      lo_table = cl_abap_tabledescr=>create( lo_struc ).

      CREATE DATA wo_data   TYPE HANDLE lo_table.
      CREATE DATA lo_data   TYPE HANDLE lo_struc.

      ASSIGN wo_data->* TO <fs_ttab>.
      ASSIGN lo_data->* TO <fs_stab>.
      LOOP AT <fs_table>  ASSIGNING <fs>.
        CLEAR <fs_stab>.
        MOVE-CORRESPONDING <fs> TO <fs_stab>.
        APPEND <fs_stab> TO <fs_ttab>.
      ENDLOOP.
    ENDIF.
  ENDIF.

  endmethod.


METHOD create_text_subtotal.
  DATA: l_string(256) TYPE c,
        l_func   TYPE string.

  CASE i_totals_function.
    WHEN zcl_excel_table=>totals_function_sum.     " Total
      l_func = 'Total'(003).
    WHEN zcl_excel_table=>totals_function_min.     " Minimum
      l_func = 'Minimum'(004).
    WHEN zcl_excel_table=>totals_function_max.     " Maximum
      l_func = 'Maximum'(005).
    WHEN zcl_excel_table=>totals_function_average. " Mean Value
      l_func = 'Average'(006).
    WHEN zcl_excel_table=>totals_function_count.   " Count
      l_func = 'Count'(007).
    WHEN OTHERS.
      CLEAR l_func.
  ENDCASE.

  MOVE i_value TO l_string.

  CONCATENATE l_string l_func INTO r_text SEPARATED BY space.

ENDMETHOD.


method CREATE_WORKSHEET.
  DATA: l_freeze_col TYPE i.

  IF wo_data IS BOUND AND wo_worksheet IS BOUND.

    wo_worksheet->zif_excel_sheet_properties~summarybelow = zif_excel_sheet_properties=>c_below_on. " By default is on

    IF wt_fieldcatalog IS INITIAL.
      set_fieldcatalog( ) .
    ELSE.
      clean_fieldcatalog( ) .
    ENDIF.

    IF i_table = abap_true.
      l_freeze_col = bind_table( i_style_table = i_style_table ) .
    ELSEIF wt_filter IS NOT INITIAL.
* Let's check for filter.
      wo_autofilter = wo_excel->add_new_autofilter( io_sheet = wo_worksheet ).
      l_freeze_col = bind_cells( ) .
      set_autofilter_area( ) .
    ELSE.
      l_freeze_col = bind_cells( ) .
    ENDIF.

* Check for freeze panes
    IF ws_layout-is_fixed = abap_true.
      IF l_freeze_col = 0.
        l_freeze_col = w_col_int.
      ENDIF.
      wo_worksheet->freeze_panes( EXPORTING ip_num_columns = l_freeze_col
                                            ip_num_rows    = w_row_int ) .
    ENDIF.
  ENDIF.

  endmethod.


method EXECUTE_CONVERTER.
  DATA: lo_if                   TYPE REF TO zif_excel_converter,
        ls_types                TYPE ts_alv_types,
        lo_addit                TYPE REF TO cl_abap_classdescr,
        lo_addit_superclass     type ref to cl_abap_classdescr.

  IF io_object IS BOUND.
    TRY.
        lo_addit            ?= cl_abap_typedescr=>describe_by_object_ref( io_object ).
      CATCH cx_sy_move_cast_error.
        RAISE EXCEPTION TYPE zcx_excel.
    ENDTRY.
    ls_types-seoclass = lo_addit->get_relative_name( ).
    READ TABLE wt_objects INTO ls_types WITH TABLE KEY seoclass = ls_types-seoclass.
    if sy-subrc ne 0.
      do.
        free lo_addit_superclass.
        lo_addit_superclass = lo_addit->get_super_class_type( ).
        if lo_addit_superclass is initial.
          sy-subrc = '4'.
          exit.
        endif.
        lo_addit = lo_addit_superclass.
        ls_types-seoclass = lo_addit->get_relative_name( ).
        read table wt_objects into ls_types with table key seoclass = ls_types-seoclass.
        if sy-subrc eq 0.
          exit.
        endif.
      enddo.
    endif.
    if sy-subrc = 0.
      CREATE OBJECT lo_if type (ls_types-clsname).
      lo_if->create_fieldcatalog(
        exporting
          is_option       = ws_option
          io_object       = io_object
          it_table        = it_table
        importing
          es_layout       = ws_layout
          et_fieldcatalog = wt_fieldcatalog
          eo_table        = wo_table
          et_colors       = wt_colors
          et_filter       = wt_filter
          ).
*  data lines of highest level.
      if ws_layout-max_subtotal_level > 0. add 1 to ws_layout-max_subtotal_level. endif.
    else.
      RAISE EXCEPTION type zcx_excel.
    endif.
  else.
    refresh wt_fieldcatalog.
    get reference of it_table into wo_table.
  endif.
endmethod.


  METHOD get_alv_converters.
    DATA:
      lt_direct_implementations TYPE seor_implementing_keys,
      lt_all_implementations    TYPE seor_implementing_keys,
      ls_impkey                 TYPE seor_implementing_key,
      ls_classkey               TYPE seoclskey,
      lr_implementation         TYPE REF TO zif_excel_converter,
      ls_object                 TYPE ts_alv_types,
      lr_classdescr             TYPE REF TO cl_abap_classdescr.

    ls_classkey-clsname = 'ZIF_EXCEL_CONVERTER'.

    CALL FUNCTION 'SEO_INTERFACE_IMPLEM_GET_ALL'
      EXPORTING
        intkey  = ls_classkey
      IMPORTING
        impkeys = lt_direct_implementations
      EXCEPTIONS
        OTHERS  = 2.

    CHECK sy-subrc = 0.

    LOOP AT lt_direct_implementations INTO ls_impkey.
      lr_classdescr ?= cl_abap_classdescr=>describe_by_name( ls_impkey-clsname ).
      IF lr_classdescr->is_instantiatable( ) = abap_true.
        APPEND ls_impkey TO lt_all_implementations.
      ENDIF.

      ls_classkey-clsname = ls_impkey-clsname.
      get_subclasses( EXPORTING is_clskey = ls_classkey CHANGING ct_classes = lt_all_implementations ).
    ENDLOOP.

    SORT lt_all_implementations BY clsname.
    DELETE ADJACENT DUPLICATES FROM lt_all_implementations COMPARING clsname.

    LOOP AT lt_all_implementations into ls_impkey.
      CLEAR ls_object.
      CREATE OBJECT lr_implementation TYPE (ls_impkey-clsname).
      ls_object-seoclass = lr_implementation->get_supported_class( ).

      IF ls_object-seoclass IS NOT INITIAL.
        ls_object-clsname  = ls_impkey-clsname.
        INSERT ls_object INTO TABLE wt_objects.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.


method GET_COLOR_STYLE.
  DATA: ls_colors         TYPE zexcel_s_converter_col,
        ls_color_styles   TYPE ts_color_styles,
        lo_style          TYPE REF TO zcl_excel_style.

  r_style = i_style. " Default we change nothing

  IF wt_colors IS NOT INITIAL.
* Full line has color
    READ TABLE wt_colors INTO ls_colors WITH KEY rownumber   = i_row
                                                 columnname  = space.
    IF sy-subrc = 0.
      READ TABLE wt_color_styles INTO ls_color_styles WITH KEY guid_old  = i_style
                                                               fontcolor = ls_colors-fontcolor
                                                               fillcolor = ls_colors-fillcolor.
      IF sy-subrc = 0.
        r_style = ls_color_styles-style_new->get_guid( ).
      ELSE.
        lo_style = create_color_style( i_style          = i_style
                                       is_colors        = ls_colors ) .
        r_style = lo_style->get_guid( ) .
        ls_color_styles-guid_old  = i_style.
        ls_color_styles-fontcolor = ls_colors-fontcolor.
        ls_color_styles-fillcolor = ls_colors-fillcolor.
        ls_color_styles-style_new = lo_style.
        INSERT ls_color_styles INTO TABLE wt_color_styles.
      ENDIF.
    ELSE.
* Only field has color
      READ TABLE wt_colors INTO ls_colors WITH KEY rownumber   = i_row
                                                   columnname  = i_fieldname.
      IF sy-subrc = 0.
        READ TABLE wt_color_styles INTO ls_color_styles WITH KEY guid_old  = i_style
                                                                 fontcolor = ls_colors-fontcolor
                                                                 fillcolor = ls_colors-fillcolor.
        IF sy-subrc = 0.
          r_style = ls_color_styles-style_new->get_guid( ).
        ELSE.
          lo_style = create_color_style( i_style          = i_style
                                         is_colors        = ls_colors ) .
          ls_color_styles-guid_old  = i_style.
          ls_color_styles-fontcolor = ls_colors-fontcolor.
          ls_color_styles-fillcolor = ls_colors-fillcolor.
          ls_color_styles-style_new = lo_style.
          INSERT ls_color_styles INTO TABLE wt_color_styles.
          r_style = ls_color_styles-style_new->get_guid( ).
        ENDIF.
      ELSE.
        r_style = i_style.
      ENDIF.
    ENDIF.
  ELSE.
    r_style = i_style.
  ENDIF.

  endmethod.


method GET_FILE.
  data: lo_excel_writer         type ref to zif_excel_writer.

  data: ls_seoclass type seoclass.


  if wo_excel is bound.
    create object lo_excel_writer type zcl_excel_writer_2007.
    e_file = lo_excel_writer->write_file( wo_excel ).

    select single * into ls_seoclass
      from seoclass
      where clsname = 'CL_BCS_CONVERT'.

    if sy-subrc = 0.
      call method (ls_seoclass-clsname)=>xstring_to_solix
        exporting
          iv_xstring = e_file
        receiving
          et_solix   = et_file.
      e_bytecount = xstrlen( e_file ).
    else.
      " Convert to binary
      call function 'SCMS_XSTRING_TO_BINARY'
        exporting
          buffer        = e_file
        importing
          output_length = e_bytecount
        tables
          binary_tab    = et_file.
    endif.
  endif.

  endmethod.


method GET_FUNCTION_NUMBER.
*Number    Function
*1  AVERAGE
*2  COUNT
*3  COUNTA
*4  MAX
*5  MIN
*6  PRODUCT
*7  STDEV
*8  STDEVP
*9  SUM
*10  VAR
*11  VARP

  case i_totals_function.
    when ZCL_EXCEL_TABLE=>TOTALS_FUNCTION_SUM.     " Total
      r_function_number = 9.
    when ZCL_EXCEL_TABLE=>TOTALS_FUNCTION_MIN.     " Minimum
      r_function_number = 5.
    when ZCL_EXCEL_TABLE=>TOTALS_FUNCTION_MAX.     " Maximum
      r_function_number = 4.
    when ZCL_EXCEL_TABLE=>TOTALS_FUNCTION_AVERAGE. " Mean Value
      r_function_number = 1.
    when ZCL_EXCEL_TABLE=>TOTALS_FUNCTION_count.   " Count
      r_function_number = 2.
    when others.
      clear r_function_number.
  endcase.
  endmethod.


method GET_OPTION.

  rs_option = ws_option.

  endmethod.


method GET_STYLE.
  DATA:   ls_styles TYPE ts_styles,
          lo_style  TYPE REF TO zcl_excel_style.

  CLEAR r_style.

  READ TABLE wt_styles INTO ls_styles WITH TABLE KEY type      = i_type
                                                     alignment = i_alignment
                                                     inttype   = i_inttype
                                                     decimals  = i_decimals.
  IF sy-subrc = 0.
    r_style = ls_styles-guid.
  ELSE.
    CASE i_type.
      WHEN c_type_hdr. " Header
        lo_style = create_style_hdr( i_alignment = i_alignment ).
      WHEN c_type_str. "Stripped
        lo_style   = create_style_stripped( i_alignment = i_alignment
                                            i_inttype   = i_inttype
                                            i_decimals  = i_decimals   ).
      WHEN c_type_nor. "Normal
        lo_style   = create_style_normal( i_alignment = i_alignment
                                          i_inttype   = i_inttype
                                          i_decimals  = i_decimals   ).
      WHEN c_type_sub. "Subtotals
        lo_style   = create_style_subtotal( i_alignment = i_alignment
                                            i_inttype   = i_inttype
                                            i_decimals  = i_decimals   ).
      WHEN c_type_tot. "Totals
        lo_style   = create_style_total( i_alignment = i_alignment
                                         i_inttype   = i_inttype
                                         i_decimals  = i_decimals   ).
    ENDCASE.
    IF lo_style IS NOT INITIAL.
      r_style = lo_style->get_guid( ).
      ls_styles-type       = i_type.
      ls_styles-alignment  = i_alignment.
      ls_styles-inttype    = i_inttype.
      ls_styles-decimals   = i_decimals.
      ls_styles-guid       = r_style.
      ls_styles-style      = lo_style.
      INSERT ls_styles INTO TABLE wt_styles.
    ENDIF.
  ENDIF.
  endmethod.


  METHOD get_subclasses.
    DATA:
      lt_subclasses TYPE seor_inheritance_keys,
      ls_subclass   TYPE seor_inheritance_key,
      lr_classdescr TYPE REF TO cl_abap_classdescr.

    CALL FUNCTION 'SEO_CLASS_GET_ALL_SUBS'
      EXPORTING
        clskey             = is_clskey
      IMPORTING
        inhkeys            = lt_subclasses
      EXCEPTIONS
        class_not_existing = 1
        OTHERS             = 2.

    CHECK sy-subrc = 0.

    LOOP AT lt_subclasses INTO ls_subclass.
      lr_classdescr ?= cl_abap_classdescr=>describe_by_name( ls_subclass-clsname ).
      IF lr_classdescr->is_instantiatable( ) = abap_true.
        APPEND ls_subclass TO ct_classes.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.


method INIT_OPTION.

  ws_option-filter = abap_true.
  ws_option-hidenc = abap_true.
  ws_option-subtot = abap_true.

  endmethod.


method LOOP_NORMAL.
  DATA: l_row_int_end     TYPE zexcel_cell_row,
        l_row_int         TYPE zexcel_cell_row,
        l_col_int         TYPE zexcel_cell_column,
        l_col_alpha       TYPE zexcel_cell_column_alpha,
        l_cell_value      TYPE zexcel_cell_value,
        l_s_color         TYPE abap_bool,
        lo_column         TYPE REF TO zcl_excel_column,
        l_formula         TYPE zexcel_cell_formula,
        l_style           TYPE zexcel_cell_style,
        l_cells           TYPE i,
        l_count           TYPE i,
        l_table_row       TYPE i.

  FIELD-SYMBOLS: <fs_stab>        TYPE ANY,
                 <fs_tab>         TYPE STANDARD TABLE,
                 <fs_sfcat>       TYPE zexcel_s_converter_fcat,
                 <fs_fldval>      TYPE ANY.

  ASSIGN wo_data->* TO <fs_tab> .

  DESCRIBE TABLE wt_fieldcatalog LINES l_cells.
  DESCRIBE TABLE <fs_tab> LINES l_count.
  l_cells = l_cells * l_count.

* It is better to loop column by column
  LOOP AT wt_fieldcatalog ASSIGNING <fs_sfcat>.
    l_row_int = i_row_int.
    l_col_int = i_col_int + <fs_sfcat>-position - 1.

*   Freeze panes
    IF <fs_sfcat>-fix_column = abap_true.
      ADD 1 TO r_freeze_col.
    ENDIF.
    l_s_color = abap_true.

    l_col_alpha = zcl_excel_common=>convert_column2alpha( l_col_int ).

*   Only if the Header is required create it.
    IF ws_option-hidehd IS INITIAL.
      " First of all write column header
      l_cell_value = <fs_sfcat>-scrtext_m.
      wo_worksheet->set_cell( ip_column    = l_col_alpha
                              ip_row       = l_row_int
                              ip_value     = l_cell_value
                              ip_style     = <fs_sfcat>-style_hdr ).
      ADD 1 TO l_row_int.
    ENDIF.
    LOOP AT <fs_tab> ASSIGNING <fs_stab>.
      l_table_row = sy-tabix.
* Now the cell values
      ASSIGN COMPONENT <fs_sfcat>-columnname OF STRUCTURE <fs_stab> TO <fs_fldval>.
* Now let's write the cell values
      IF ws_layout-is_stripped = abap_true AND l_s_color = abap_true.
        l_style = get_color_style( i_row       = l_table_row
                                   i_fieldname = <fs_sfcat>-columnname
                                   i_style     = <fs_sfcat>-style_stripped  ).
        wo_worksheet->set_cell( ip_column    = l_col_alpha
                                ip_row       = l_row_int
                                ip_value     = <fs_fldval>
                                ip_style     = l_style ).
        CLEAR l_s_color.
      ELSE.
        l_style = get_color_style( i_row       = l_table_row
                                   i_fieldname = <fs_sfcat>-columnname
                                   i_style     = <fs_sfcat>-style_normal  ).
        wo_worksheet->set_cell( ip_column    = l_col_alpha
                                ip_row       = l_row_int
                                ip_value     = <fs_fldval>
                                ip_style     = l_style  ).
        l_s_color = abap_true.
      ENDIF.
      READ TABLE wt_filter TRANSPORTING NO FIELDS WITH TABLE KEY rownumber  = l_table_row
                                                                 columnname = <fs_sfcat>-columnname.
      IF sy-subrc = 0.
        wo_worksheet->get_cell( EXPORTING
                                   ip_column    = l_col_alpha
                                   ip_row       = l_row_int
                                IMPORTING
                                   ep_value     = l_cell_value ).
        wo_autofilter->set_value( i_column = l_col_int
                                  i_value  = l_cell_value ).
      ENDIF.
      ADD 1 TO l_row_int.
    ENDLOOP.
* Now let's check for optimized
    IF <fs_sfcat>-is_optimized = abap_true .
      lo_column = wo_worksheet->get_column( ip_column = l_col_alpha ).
      lo_column->set_auto_size( ip_auto_size = abap_true ) .
    ENDIF.
* Now let's check for visible
    IF <fs_sfcat>-is_hidden = abap_true.
      lo_column = wo_worksheet->get_column( ip_column = l_col_alpha ).
      lo_column->set_visible( ip_visible = abap_false ) .
    ENDIF.
* Now let's check for total versus subtotal.
    IF <fs_sfcat>-totals_function IS NOT INITIAL.
      l_row_int_end = l_row_int - 1.

      l_formula = create_formular_total( i_row_int         = l_row_int_end
                                         i_column          = l_col_alpha
                                         i_totals_function = <fs_sfcat>-totals_function ).
      wo_worksheet->set_cell( ip_column    = l_col_alpha
                              ip_row       = l_row_int
                              ip_formula   = l_formula
                              ip_style     = <fs_sfcat>-style_total ).
    ENDIF.
  ENDLOOP.
  endmethod.


method LOOP_SUBTOTAL.

  DATA: l_row_int_start   TYPE zexcel_cell_row,
        l_row_int_end     TYPE zexcel_cell_row,
        l_row_int         TYPE zexcel_cell_row,
        l_col_int         TYPE zexcel_cell_column,
        l_col_alpha       TYPE zexcel_cell_column_alpha,
        l_col_alpha_start TYPE zexcel_cell_column_alpha,
        l_cell_value      TYPE zexcel_cell_value,
        l_s_color         TYPE abap_bool,
        lo_column         TYPE REF TO zcl_excel_column,
        lo_row            TYPE REF TO zcl_excel_row,
        l_formula         TYPE zexcel_cell_formula,
        l_style           TYPE zexcel_cell_style,
        l_text            TYPE string,
        ls_sort_values    TYPE ts_sort_values,
        ls_subtotal_rows  TYPE ts_subtotal_rows,
        l_sort_level      TYPE int4,
        l_hidden          TYPE int4,
        l_line            TYPE i,
        l_cells           TYPE i,
        l_count           TYPE i,
        l_table_row       TYPE i,
        lt_fcat           TYPE zexcel_t_converter_fcat.

  FIELD-SYMBOLS: <fs_stab>        TYPE ANY,
                 <fs_tab>         TYPE STANDARD TABLE,
                 <fs_sfcat>       TYPE zexcel_s_converter_fcat,
                 <fs_fldval>      TYPE ANY,
                 <fs_sortval>     TYPE ANY,
                 <fs_sortv>       TYPE ts_sort_values.

  ASSIGN wo_data->* TO <fs_tab> .

  REFRESH: wt_sort_values,
           wt_subtotal_rows.

  DESCRIBE TABLE wt_fieldcatalog LINES l_cells.
  DESCRIBE TABLE <fs_tab> LINES l_count.
  l_cells = l_cells * l_count.

  READ TABLE <fs_tab> ASSIGNING <fs_stab> INDEX 1.
  IF sy-subrc = 0.
    l_row_int = i_row_int + 1.
    lt_fcat =  wt_fieldcatalog.
    SORT lt_fcat BY sort_level DESCENDING.
    LOOP AT lt_fcat ASSIGNING <fs_sfcat> WHERE is_subtotalled = abap_true.
      ASSIGN COMPONENT <fs_sfcat>-columnname OF STRUCTURE <fs_stab> TO <fs_fldval>.
      ls_sort_values-fieldname    = <fs_sfcat>-columnname.
      ls_sort_values-row_int      = l_row_int.
      ls_sort_values-sort_level   = <fs_sfcat>-sort_level.
      ls_sort_values-is_collapsed = <fs_sfcat>-is_collapsed.
      CREATE DATA ls_sort_values-value LIKE <fs_fldval>.
      ASSIGN ls_sort_values-value->* TO <fs_sortval>.
      <fs_sortval> =  <fs_fldval>.
      INSERT ls_sort_values INTO TABLE wt_sort_values.
    ENDLOOP.
  ENDIF.
  l_row_int = i_row_int.
* Let's check if we need to hide a sort level.
  DESCRIBE TABLE wt_sort_values LINES l_line.
  IF  l_line <= 1.
    CLEAR l_hidden.
  ELSE.
    LOOP AT wt_sort_values INTO ls_sort_values WHERE is_collapsed = abap_false.
      IF l_hidden < ls_sort_values-sort_level.
        l_hidden = ls_sort_values-sort_level.
      ENDIF.
    ENDLOOP.
  ENDIF.
  ADD 1 TO l_hidden. " As this is the first level we show.
* First loop without formular only addtional rows with subtotal text.
  LOOP AT <fs_tab> ASSIGNING <fs_stab>.
    ADD 1 TO l_row_int.  " 1 is for header row.
    l_row_int_start = l_row_int.
    SORT lt_fcat BY sort_level DESCENDING.
    LOOP AT lt_fcat ASSIGNING <fs_sfcat> WHERE is_subtotalled = abap_true.
      l_col_int   = i_col_int + <fs_sfcat>-position - 1.
      l_col_alpha = zcl_excel_common=>convert_column2alpha( l_col_int ).
* Now the cell values
      ASSIGN COMPONENT <fs_sfcat>-columnname OF STRUCTURE <fs_stab> TO <fs_fldval>.
      IF sy-subrc = 0.
        READ TABLE wt_sort_values ASSIGNING <fs_sortv> WITH TABLE KEY fieldname = <fs_sfcat>-columnname.
        IF sy-subrc = 0.
          ASSIGN <fs_sortv>-value->* TO <fs_sortval>.
          IF <fs_sortval> <> <fs_fldval> OR <fs_sortv>-new = abap_true.
* First let's remmember the subtotal values as it has to appear later.
            ls_subtotal_rows-row_int       = l_row_int.
            ls_subtotal_rows-row_int_start = <fs_sortv>-row_int.
            ls_subtotal_rows-columnname    = <fs_sfcat>-columnname.
            INSERT ls_subtotal_rows INTO TABLE wt_subtotal_rows.
* Now let's write the subtotal line
            l_cell_value = create_text_subtotal( i_value = <fs_sortval>
                                   i_totals_function     = <fs_sfcat>-totals_function ).
            wo_worksheet->set_cell( ip_column    = l_col_alpha
                                    ip_row       = l_row_int
                                    ip_value     = l_cell_value
                                    ip_abap_type = cl_abap_typedescr=>typekind_string
                                    ip_style     = <fs_sfcat>-style_subtotal  ).
            lo_row = wo_worksheet->get_row( ip_row = l_row_int ).
            lo_row->set_outline_level( ip_outline_level = <fs_sfcat>-sort_level ) .
            IF <fs_sfcat>-is_collapsed = abap_true.
              IF <fs_sfcat>-sort_level >  l_hidden.
                lo_row->set_visible( ip_visible =  abap_false ) .
              ENDIF.
              lo_row->set_collapsed( ip_collapsed =  <fs_sfcat>-is_collapsed ) .
            ENDIF.
* Now let's change the key
            ADD 1 TO l_row_int.
            <fs_sortval> =  <fs_fldval>.
            <fs_sortv>-new = abap_false.
            l_line = <fs_sortv>-sort_level.
            LOOP AT wt_sort_values ASSIGNING <fs_sortv> WHERE sort_level >= l_line.
              <fs_sortv>-row_int = l_row_int.
            ENDLOOP.
          ENDIF.
        ENDIF.
      ENDIF.
    ENDLOOP.
  ENDLOOP.
  ADD 1 TO l_row_int.
  l_row_int_start = l_row_int.
  SORT lt_fcat BY sort_level DESCENDING.
  LOOP AT lt_fcat ASSIGNING <fs_sfcat> WHERE is_subtotalled = abap_true.
    l_col_int   = i_col_int + <fs_sfcat>-position - 1.
    l_col_alpha = zcl_excel_common=>convert_column2alpha( l_col_int ).
    READ TABLE wt_sort_values ASSIGNING <fs_sortv> WITH TABLE KEY fieldname = <fs_sfcat>-columnname.
    IF sy-subrc = 0.
      ASSIGN <fs_sortv>-value->* TO <fs_sortval>.
      ls_subtotal_rows-row_int       = l_row_int.
      ls_subtotal_rows-row_int_start = <fs_sortv>-row_int.
      ls_subtotal_rows-columnname    = <fs_sfcat>-columnname.
      INSERT ls_subtotal_rows INTO TABLE wt_subtotal_rows.
* First let's write the value as it has to appear.
      l_cell_value = create_text_subtotal( i_value = <fs_sortval>
                             i_totals_function     = <fs_sfcat>-totals_function ).
      wo_worksheet->set_cell( ip_column    = l_col_alpha
                              ip_row       = l_row_int
                              ip_value     = l_cell_value
                              ip_abap_type = cl_abap_typedescr=>typekind_string
                              ip_style     = <fs_sfcat>-style_subtotal ).

      l_sort_level = <fs_sfcat>-sort_level.
      lo_row = wo_worksheet->get_row( ip_row = l_row_int ).
      lo_row->set_outline_level( ip_outline_level = l_sort_level ) .
      IF <fs_sfcat>-is_collapsed = abap_true.
        IF <fs_sfcat>-sort_level >  l_hidden.
          lo_row->set_visible( ip_visible =  abap_false ) .
        ENDIF.
        lo_row->set_collapsed( ip_collapsed =  <fs_sfcat>-is_collapsed ) .
      ENDIF.
      ADD 1 TO l_row_int.
    ENDIF.
  ENDLOOP.
* Let's write the Grand total
  l_sort_level = 0.
  lo_row = wo_worksheet->get_row( ip_row = l_row_int ).
  lo_row->set_outline_level( ip_outline_level = l_sort_level ) .
*  lo_row_dim->set_collapsed( ip_collapsed =  <fs_sfcat>-is_collapsed ) . Not on grand total

  l_text    = create_text_subtotal( i_value = 'Grand'(002)
                                    i_totals_function     = <fs_sfcat>-totals_function ).

  l_col_alpha_start = zcl_excel_common=>convert_column2alpha( i_col_int ).
  wo_worksheet->set_cell( ip_column    = l_col_alpha_start
                          ip_row       = l_row_int
                          ip_value     = l_text
                          ip_abap_type = cl_abap_typedescr=>typekind_string
                          ip_style     = <fs_sfcat>-style_subtotal ).

* It is better to loop column by column second time around
* Second loop with formular and data.
  LOOP AT wt_fieldcatalog ASSIGNING <fs_sfcat>.
    l_row_int = i_row_int.
    l_col_int = i_col_int + <fs_sfcat>-position - 1.
* Freeze panes
    IF <fs_sfcat>-fix_column = abap_true.
      ADD 1 TO r_freeze_col.
    ENDIF.
    l_s_color = abap_true.
    l_col_alpha = zcl_excel_common=>convert_column2alpha( l_col_int ).
    " First of all write column header
    l_cell_value = <fs_sfcat>-scrtext_m.
    wo_worksheet->set_cell( ip_column    = l_col_alpha
                            ip_row       = l_row_int
                            ip_value     = l_cell_value
                            ip_abap_type = cl_abap_typedescr=>typekind_string
                            ip_style     = <fs_sfcat>-style_hdr ).
    ADD 1 TO l_row_int.
    LOOP AT <fs_tab> ASSIGNING <fs_stab>.
      l_table_row = sy-tabix.
* Now the cell values
      ASSIGN COMPONENT <fs_sfcat>-columnname OF STRUCTURE <fs_stab> TO <fs_fldval>.
* Let's check for subtotal lines
      DO.
        READ TABLE wt_subtotal_rows TRANSPORTING NO FIELDS WITH TABLE KEY row_int = l_row_int.
        IF sy-subrc = 0.
          IF <fs_sfcat>-is_subtotalled = abap_false AND
             <fs_sfcat>-totals_function IS NOT INITIAL.
            DO.
              READ TABLE wt_subtotal_rows INTO ls_subtotal_rows WITH TABLE KEY row_int    = l_row_int.
              IF sy-subrc = 0.
                l_row_int_start = ls_subtotal_rows-row_int_start.
                l_row_int_end   = l_row_int - 1.

                l_formula = create_formular_subtotal( i_row_int_start   = l_row_int_start
                                                      i_row_int_end     = l_row_int_end
                                                      i_column          = l_col_alpha
                                                      i_totals_function = <fs_sfcat>-totals_function ).
                wo_worksheet->set_cell( ip_column    = l_col_alpha
                                        ip_row       = l_row_int
                                        ip_formula   = l_formula
                                        ip_style     = <fs_sfcat>-style_subtotal ).
                IF <fs_sfcat>-is_collapsed = abap_true.
                  lo_row = wo_worksheet->get_row( ip_row = l_row_int ).
                  lo_row->set_collapsed( ip_collapsed =  <fs_sfcat>-is_collapsed ).
                  IF <fs_sfcat>-sort_level >  l_hidden.
                    lo_row->set_visible( ip_visible =  abap_false ) .
                  ENDIF.
                ENDIF.
                ADD 1 TO l_row_int.
              ELSE.
                EXIT.
              ENDIF.
            ENDDO.
          ELSE.
            ADD 1 TO l_row_int.
          ENDIF.
        ELSE.
          EXIT.
        ENDIF.
      ENDDO.
* Let's set the row dimension values
      lo_row = wo_worksheet->get_row( ip_row = l_row_int ).
      lo_row->set_outline_level( ip_outline_level = ws_layout-max_subtotal_level ) .
      IF <fs_sfcat>-is_collapsed  = abap_true.
        lo_row->set_visible( ip_visible =  abap_false ) .
        lo_row->set_collapsed( ip_collapsed =  <fs_sfcat>-is_collapsed ) .
      ENDIF.
* Now let's write the cell values
      IF ws_layout-is_stripped = abap_true AND l_s_color = abap_true.
        l_style = get_color_style( i_row       = l_table_row
                                   i_fieldname = <fs_sfcat>-columnname
                                   i_style     = <fs_sfcat>-style_stripped  ).
        wo_worksheet->set_cell( ip_column    = l_col_alpha
                                ip_row       = l_row_int
                                ip_value     = <fs_fldval>
                                ip_style     = l_style ).
        CLEAR l_s_color.
      ELSE.
        l_style = get_color_style( i_row       = l_table_row
                                   i_fieldname = <fs_sfcat>-columnname
                                   i_style     = <fs_sfcat>-style_normal  ).
        wo_worksheet->set_cell( ip_column    = l_col_alpha
                                ip_row       = l_row_int
                                ip_value     = <fs_fldval>
                                ip_style     = l_style   ).
        l_s_color = abap_true.
      ENDIF.
      READ TABLE wt_filter TRANSPORTING NO FIELDS WITH TABLE KEY rownumber  = l_table_row
                                                                 columnname = <fs_sfcat>-columnname.
      IF sy-subrc = 0.
        wo_worksheet->get_cell( EXPORTING
                                   ip_column    = l_col_alpha
                                   ip_row       = l_row_int
                                IMPORTING
                                   ep_value     = l_cell_value ).
        wo_autofilter->set_value( i_column = l_col_int
                                  i_value  = l_cell_value ).
      ENDIF.
      ADD 1 TO l_row_int.
    ENDLOOP.
* Let's check for subtotal lines
    DO.
      READ TABLE wt_subtotal_rows TRANSPORTING NO FIELDS WITH TABLE KEY row_int = l_row_int.
      IF sy-subrc = 0.
        IF <fs_sfcat>-is_subtotalled = abap_false AND
           <fs_sfcat>-totals_function IS NOT INITIAL.
          DO.
            READ TABLE wt_subtotal_rows INTO ls_subtotal_rows WITH TABLE KEY row_int    = l_row_int.
            IF sy-subrc = 0.
              l_row_int_start = ls_subtotal_rows-row_int_start.
              l_row_int_end   = l_row_int - 1.

              l_formula = create_formular_subtotal( i_row_int_start   = l_row_int_start
                                                    i_row_int_end     = l_row_int_end
                                                    i_column          = l_col_alpha
                                                    i_totals_function = <fs_sfcat>-totals_function ).
              wo_worksheet->set_cell( ip_column    = l_col_alpha
                                      ip_row       = l_row_int
                                      ip_formula   = l_formula
                                      ip_style     = <fs_sfcat>-style_subtotal ).
              IF <fs_sfcat>-is_collapsed = abap_true.
                lo_row = wo_worksheet->get_row( ip_row = l_row_int ).
                lo_row->set_collapsed( ip_collapsed =  <fs_sfcat>-is_collapsed ).
              ENDIF.
              ADD 1 TO l_row_int.
            ELSE.
              EXIT.
            ENDIF.
          ENDDO.
        ELSE.
          ADD 1 TO l_row_int.
        ENDIF.
      ELSE.
        EXIT.
      ENDIF.
    ENDDO.
* Now let's check for Grand total
    IF <fs_sfcat>-is_subtotalled = abap_false AND
       <fs_sfcat>-totals_function IS NOT INITIAL.
      l_row_int_start = i_row_int + 1.
      l_row_int_end   = l_row_int - 1.

      l_formula = create_formular_subtotal( i_row_int_start   = l_row_int_start
                                            i_row_int_end     = l_row_int_end
                                            i_column          = l_col_alpha
                                            i_totals_function = <fs_sfcat>-totals_function ).
      wo_worksheet->set_cell( ip_column    = l_col_alpha
                              ip_row       = l_row_int
                              ip_formula   = l_formula
                              ip_style     = <fs_sfcat>-style_subtotal ).
    ENDIF.
* Now let's check for optimized
    IF <fs_sfcat>-is_optimized = abap_true.
      lo_column = wo_worksheet->get_column( ip_column = l_col_alpha ).
      lo_column->set_auto_size( ip_auto_size = abap_true ) .
    ENDIF.
* Now let's check for visible
    IF <fs_sfcat>-is_hidden = abap_true.
      lo_column = wo_worksheet->get_column( ip_column = l_col_alpha ).
      lo_column->set_visible( ip_visible = abap_false ) .
    ENDIF.
  ENDLOOP.

  endmethod.


method OPEN_FILE.
  data: l_bytecount             type i,
        lt_file                 type solix_tab,
        l_dir                   type string.

  field-symbols: <fs_data> type any table.

  assign wo_data->* to <fs_data>.

* catch zcx_excel .
*endtry.
  if wo_excel is bound.
    get_file( importing e_bytecount  = l_bytecount
                        et_file      = lt_file ) .

    l_dir =  create_path( ) .

    cl_gui_frontend_services=>gui_download( exporting bin_filesize = l_bytecount
                                                      filename     = l_dir
                                                      filetype     = 'BIN'
                                             changing data_tab     = lt_file ).
    cl_gui_frontend_services=>execute(
      exporting
        document               = l_dir
*        application            =
*        parameter              =
*        default_directory      =
*        maximized              =
*        minimized              =
*        synchronous            =
*        operation              = 'OPEN'
      exceptions
        cntl_error             = 1
        error_no_gui           = 2
        bad_parameter          = 3
        file_not_found         = 4
        path_not_found         = 5
        file_extension_unknown = 6
        error_execute_failed   = 7
        synchronous_failed     = 8
        not_supported_by_gui   = 9
           ).
    if sy-subrc <> 0.
      message id sy-msgid type sy-msgty number sy-msgno
                 with sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    endif.

  endif.


  endmethod.


method SET_AUTOFILTER_AREA.
  DATA: ls_area TYPE zexcel_s_autofilter_area,
        l_lines TYPE i,
        lt_values TYPE zexcel_t_autofilter_values,
        ls_values TYPE zexcel_s_autofilter_values.

* Let's check for filter.
  IF wo_autofilter IS BOUND.
    ls_area-row_start = 1.
    lt_values = wo_autofilter->get_values( ) .
    SORT lt_values BY column ASCENDING.
    DESCRIBE TABLE lt_values LINES l_lines.
    READ TABLE lt_values INTO ls_values INDEX 1.
    IF sy-subrc = 0.
      ls_area-col_start = ls_values-column.
    ENDIF.
    READ TABLE lt_values INTO ls_values INDEX l_lines.
    IF sy-subrc = 0.
      ls_area-col_end = ls_values-column.
    ENDIF.
    wo_autofilter->set_filter_area( is_area = ls_area ) .
  ENDIF.

  endmethod.


method SET_CELL_FORMAT.
  DATA:  l_format    TYPE zexcel_number_format.

  CLEAR r_format.
  CASE i_inttype.
    WHEN cl_abap_typedescr=>typekind_date.
      r_format = wo_worksheet->get_default_excel_date_format( ).
    WHEN cl_abap_typedescr=>typekind_time.
      r_format = wo_worksheet->get_default_excel_time_format( ).
    WHEN cl_abap_typedescr=>typekind_float OR cl_abap_typedescr=>typekind_packed.
      IF i_decimals > 0 .
        l_format = '#,##0.'.
        DO i_decimals TIMES.
          CONCATENATE l_format '0' INTO l_format.
        ENDDO.
        r_format = l_format.
      ENDIF.
    WHEN cl_abap_typedescr=>typekind_int OR cl_abap_typedescr=>typekind_int1 OR cl_abap_typedescr=>typekind_int2.
      r_format = '#,##0'.
  ENDCASE.

  endmethod.


method SET_FIELDCATALOG.

  DATA: lr_data             TYPE REF TO data,
        lo_structdescr      TYPE REF TO cl_abap_structdescr,
        lt_dfies            TYPE ddfields,
        ls_dfies            TYPE dfies.
  DATA: ls_fcat             TYPE zexcel_s_converter_fcat.

  FIELD-SYMBOLS: <fs_tab>         TYPE ANY TABLE.

  ASSIGN wo_data->* TO <fs_tab> .

  CREATE DATA lr_data LIKE LINE OF <fs_tab>.

  lo_structdescr ?= cl_abap_structdescr=>describe_by_data_ref( lr_data ).

  lt_dfies = zcl_excel_common=>describe_structure( io_struct = lo_structdescr ).

  LOOP AT lt_dfies INTO ls_dfies.
    MOVE-CORRESPONDING ls_dfies TO ls_fcat.
    ls_fcat-columnname = ls_dfies-fieldname.
    INSERT ls_fcat INTO TABLE wt_fieldcatalog.
  ENDLOOP.

  clean_fieldcatalog( ).

  endmethod.


method SET_OPTION.

  IF ws_indx-begdt IS INITIAL.
    ws_indx-begdt = sy-datum.
  ENDIF.

  ws_indx-aedat = sy-datum.
  ws_indx-usera = sy-uname.
  ws_indx-pgmid = sy-cprog.

  EXPORT p1 = is_option TO DATABASE indx(xl) FROM ws_indx ID ws_indx-srtfd.

  IF sy-subrc = 0.
    ws_option = is_option.
  ENDIF.

  endmethod.


method WRITE_FILE.
  data: l_bytecount             type i,
      lt_file                 type solix_tab,
      l_dir                   type string.

  field-symbols: <fs_data> type any table.

  assign wo_data->* to <fs_data>.

* catch zcx_excel .
*endtry.
  if wo_excel is bound.
    get_file( importing e_bytecount  = l_bytecount
                        et_file      = lt_file ) .
    if i_path is initial.
      l_dir =  create_path( ) .
    else.
      l_dir = i_path.
    endif.
    cl_gui_frontend_services=>gui_download( exporting bin_filesize = l_bytecount
                                                      filename     = l_dir
                                                      filetype     = 'BIN'
                                             changing data_tab     = lt_file ).
  endif.
  endmethod.
ENDCLASS.
