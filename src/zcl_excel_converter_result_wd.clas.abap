class ZCL_EXCEL_CONVERTER_RESULT_WD definition
  public
  inheriting from ZCL_EXCEL_CONVERTER_RESULT
  final
  create public .

*"* public components of class ZCL_EXCEL_CONVERTER_RESULT_WD
*"* do not include other source files here!!!
public section.

  methods ZIF_EXCEL_CONVERTER~CAN_CONVERT_OBJECT
    redefinition .
  methods ZIF_EXCEL_CONVERTER~CREATE_FIELDCATALOG
    redefinition .
*"* protected components of class ZCL_EXCEL_CONVERTER_RESULT_WD
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_CONVERTER_RESULT_WD
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_CONVERTER_RESULT_WD
*"* do not include other source files here!!!
protected section.
private section.

  data WO_CONFIG type ref to CL_SALV_WD_CONFIG_TABLE .
  data WT_FIELDS type SALV_WD_T_FIELD_REF .
  data WT_COLUMNS type SALV_WD_T_COLUMN_REF .

  methods GET_COLUMNS_INFO
    changing
      !XS_FCAT type LVC_S_FCAT .
  methods GET_FIELDS_INFO
    changing
      !XS_FCAT type LVC_S_FCAT .
  methods CREATE_WT_SORT .
  methods CREATE_WT_FILT .
  methods CREATE_WT_FCAT
    importing
      !IO_TABLE type ref to DATA .
ENDCLASS.



CLASS ZCL_EXCEL_CONVERTER_RESULT_WD IMPLEMENTATION.


method CREATE_WT_FCAT.
  DATA: lr_data             TYPE REF TO data,
        lo_structdescr      TYPE REF TO cl_abap_structdescr,
        lt_dfies            TYPE ddfields,
        ls_dfies            TYPE dfies.

  DATA: ls_fcat         TYPE lvc_s_fcat.

  FIELD-SYMBOLS: <fs_tab>         TYPE ANY TABLE.

  ASSIGN io_table->* TO <fs_tab> .
  CREATE DATA lr_data LIKE LINE OF <fs_tab>.

  lo_structdescr ?= cl_abap_structdescr=>describe_by_data_ref( lr_data ).

  lt_dfies = zcl_excel_common=>describe_structure( io_struct = lo_structdescr ).

  loop at lt_dfies into ls_dfies.
      MOVE-CORRESPONDING ls_dfies TO ls_fcat.
*      ls_fcat-columnname = ls_dfies-fieldname.
      ls_fcat-col_pos = ls_dfies-position.
      ls_fcat-key     = ls_dfies-keyflag.
      get_fields_info( CHANGING xs_fcat = ls_fcat ) .

      ls_fcat-col_opt = abap_true.

      get_columns_info( CHANGING xs_fcat = ls_fcat ) .

      INSERT ls_fcat INTO TABLE wt_fcat.
  endloop.

  endmethod.


method CREATE_WT_FILT.
* No neeed for superclass.
* Only for WD
  DATA: lt_filters      TYPE salv_wd_t_filter_rule_ref,
        ls_filt         TYPE lvc_s_filt.

  FIELD-SYMBOLS: <fs_fields>  TYPE salv_wd_s_field_ref,
                 <fs_filter>  TYPE salv_wd_s_filter_rule_ref.

  LOOP AT  wt_fields ASSIGNING <fs_fields>.
    REFRESH lt_filters.
    lt_filters    = <fs_fields>-r_field->if_salv_wd_filter~get_filter_rules( ) .
    LOOP AT lt_filters ASSIGNING <fs_filter>.
      ls_filt-fieldname = <fs_fields>-fieldname.
      IF <fs_filter>-r_filter_rule->get_included( ) = abap_true.
        ls_filt-sign      = 'I'.
      ELSE.
        ls_filt-sign      = 'E'.
      ENDIF.
      ls_filt-option    = <fs_filter>-r_filter_rule->get_operator( ).
      ls_filt-high      = <fs_filter>-r_filter_rule->get_high_value( ) .
      ls_filt-low       = <fs_filter>-r_filter_rule->get_low_value( ) .
      INSERT ls_filt INTO TABLE wt_filt.
    ENDLOOP.
  ENDLOOP.

  endmethod.


method CREATE_WT_SORT.
  DATA: lo_sort         TYPE REF TO cl_salv_wd_sort_rule,
        l_sort_order    TYPE salv_wd_constant,
        ls_sort         TYPE lvc_s_sort.

  FIELD-SYMBOLS: <fs_fields>  TYPE salv_wd_s_field_ref.

  LOOP AT  wt_fields ASSIGNING <fs_fields>.
    lo_sort      = <fs_fields>-r_field->if_salv_wd_sort~get_sort_rule( ) .
    IF lo_sort IS BOUND.
      l_sort_order = lo_sort->get_sort_order( ).
      IF l_sort_order <> if_salv_wd_c_sort=>sort_order.
        CLEAR ls_sort.
        ls_sort-spos      = lo_sort->get_sort_position( ).
        ls_sort-fieldname = <fs_fields>-fieldname.
*      ls_sort-GROUP
        ls_sort-subtot    = lo_sort->get_group_aggregation( ).
*      ls_sort-COMP
*      ls_sort-EXPA
*      ls_sort-SELTEXT
*      ls_sort-OBLIGATORY
*      ls_sort-LEVEL
*      ls_sort-NO_OUT
        IF l_sort_order = if_salv_wd_c_sort=>sort_order_ascending.
          ls_sort-up = abap_true.
        ELSE.
          ls_sort-down = abap_true.
        ENDIF.
        INSERT ls_sort INTO TABLE wt_sort.
      ENDIF.
    ENDIF.
  ENDLOOP.

  endmethod.


method GET_COLUMNS_INFO.
  DATA:  l_numc2             TYPE salv_wd_constant.


  FIELD-SYMBOLS: <fs_column>  TYPE salv_wd_s_column_ref.

  READ TABLE wt_columns ASSIGNING <fs_column> WITH KEY id = xs_fcat-fieldname .
  IF sy-subrc = 0.
    xs_fcat-col_pos    = <fs_column>-r_column->get_position( ) .
    l_numc2 = <fs_column>-r_column->get_fixed_position( ).
    IF l_numc2 = '02'.
      xs_fcat-fix_column = abap_true .
    ENDIF.
    l_numc2 = <fs_column>-r_column->get_visible( ).
    IF l_numc2 = '01'.
      xs_fcat-no_out = abap_true .
    ENDIF.
  ENDIF.

  endmethod.


method GET_FIELDS_INFO.
  DATA:   lo_aggr         TYPE REF TO cl_salv_wd_aggr_rule,
          l_aggrtype      TYPE salv_wd_constant.

  FIELD-SYMBOLS: <fs_fields>  TYPE salv_wd_s_field_ref.

  READ TABLE wt_fields ASSIGNING <fs_fields> WITH KEY fieldname = xs_fcat-fieldname.
  IF sy-subrc = 0.
    lo_aggr = <fs_fields>-r_field->if_salv_wd_aggr~get_aggr_rule( ) .
    IF lo_aggr IS BOUND.
      l_aggrtype = lo_aggr->get_aggregation_type( ) .
      CASE l_aggrtype.
        WHEN if_salv_wd_c_aggregation=>aggrtype_total.
          xs_fcat-do_sum = abap_true.
        WHEN if_salv_wd_c_aggregation=>aggrtype_minimum.
          xs_fcat-do_sum =  'A'.
        WHEN if_salv_wd_c_aggregation=>aggrtype_maximum .
          xs_fcat-do_sum =  'B'.
        WHEN if_salv_wd_c_aggregation=>aggrtype_average .
          xs_fcat-do_sum =  'C'.
        WHEN OTHERS.
          CLEAR xs_fcat-do_sum .
      ENDCASE.
    ENDIF.
  ENDIF.

  endmethod.


METHOD ZIF_EXCEL_CONVERTER~CAN_CONVERT_OBJECT.

  DATA: lo_result TYPE REF TO cl_salv_wd_result_data_table.

  TRY.
      lo_result ?= io_object.
    CATCH cx_sy_move_cast_error .
      RAISE EXCEPTION TYPE zcx_excel.
  ENDTRY.

ENDMETHOD.


METHOD zif_excel_converter~create_fieldcatalog.
  DATA: lo_result  TYPE REF TO cl_salv_wd_result_data_table,
        lo_data    TYPE REF TO data.

  FIELD-SYMBOLS: <fs_table> TYPE STANDARD TABLE.

  zif_excel_converter~can_convert_object( io_object = io_object ).

  ws_option = is_option.

  lo_result ?= io_object.

  CLEAR: es_layout,
         et_fieldcatalog.

  IF lo_result IS BOUND.
    lo_data = get_table( io_object = lo_result->r_model->r_data ).
    IF lo_data IS BOUND.
      ASSIGN lo_data->* TO <fs_table> .

      wo_config ?= lo_result->r_model->r_model.

      IF wo_config IS BOUND.
        wt_fields  = wo_config->if_salv_wd_field_settings~get_fields( ) .
        wt_columns = wo_config->if_salv_wd_column_settings~get_columns( ) .
      ENDIF.

      create_wt_fcat( io_table = lo_data ).
      create_wt_sort( ).
      create_wt_filt( ).

      apply_sort( EXPORTING it_table = <fs_table>
                  IMPORTING eo_table = eo_table ) .

*      get_color( EXPORTING io_table    = eo_table
*                 IMPORTING et_colors   = et_colors ) .

      get_filter( IMPORTING et_filter  = et_filter
                  CHANGING  xo_table   = eo_table ) .

      update_catalog( CHANGING  cs_layout       = es_layout
                                ct_fieldcatalog = et_fieldcatalog ).
    ELSE.
* We have a problem and should stop here
    ENDIF.
  ENDIF.
ENDMETHOD.
ENDCLASS.
