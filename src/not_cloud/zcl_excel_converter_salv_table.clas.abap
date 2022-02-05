CLASS zcl_excel_converter_salv_table DEFINITION
  PUBLIC
  INHERITING FROM zcl_excel_converter_alv
  FINAL
  CREATE PUBLIC .

*"* public components of class ZCL_EXCEL_CONVERTER_SALV_TABLE
*"* do not include other source files here!!!
  PUBLIC SECTION.

    METHODS zif_excel_converter~can_convert_object
        REDEFINITION .
    METHODS zif_excel_converter~create_fieldcatalog
        REDEFINITION .
    METHODS zif_excel_converter~get_supported_class
        REDEFINITION .
*"* protected components of class ZCL_EXCEL_CONVERTER_SALV_TABLE
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_CONVERTER_SALV_TABLE
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_CONVERTER_SALV_TABLE
*"* do not include other source files here!!!
  PROTECTED SECTION.
  PRIVATE SECTION.

    METHODS load_data
      IMPORTING
        !io_salv  TYPE REF TO cl_salv_table
        !it_table TYPE STANDARD TABLE .
    METHODS is_intercept_data_active
      RETURNING
        VALUE(rv_result) TYPE abap_bool.
ENDCLASS.



CLASS zcl_excel_converter_salv_table IMPLEMENTATION.


  METHOD load_data.
    DATA: lo_columns      TYPE REF TO cl_salv_columns_table,
          lo_aggregations TYPE REF TO cl_salv_aggregations,
          lo_sorts        TYPE REF TO cl_salv_sorts,
          lo_filters      TYPE REF TO cl_salv_filters,
          lo_functional   TYPE REF TO cl_salv_functional_settings,
          lo_display      TYPE REF TO cl_salv_display_settings.

    DATA: ls_vari   TYPE disvariant,
          lo_layout TYPE REF TO cl_salv_layout.

    DATA lt_kkblo_fieldcat TYPE kkblo_t_fieldcat.
    DATA ls_kkblo_layout   TYPE kkblo_layout.
    DATA lt_kkblo_filter   TYPE kkblo_t_filter.
    DATA lt_kkblo_sort     TYPE kkblo_t_sortinfo.
    DATA: lv_intercept_data_active TYPE abap_bool,
          ls_layout_key            TYPE salv_s_layout_key.

    lo_layout               = io_salv->get_layout( ) .
    lo_columns              = io_salv->get_columns( ).
    lo_aggregations         = io_salv->get_aggregations( ) .
    lo_sorts                = io_salv->get_sorts( ) .
    lo_filters              = io_salv->get_filters( ) .
    lo_display              = io_salv->get_display_settings( ) .
    lo_functional           = io_salv->get_functional_settings( ) .

    CLEAR: wt_fcat, wt_sort, wt_filt.

    lv_intercept_data_active = is_intercept_data_active( ).

* First update metadata if we can.
    IF io_salv->is_offline( ) = abap_false.
      IF lv_intercept_data_active = abap_true.
        ls_layout_key = lo_layout->get_key( ).
        ls_vari-report    = ls_layout_key-report.
        ls_vari-handle    = ls_layout_key-handle.
        ls_vari-log_group = ls_layout_key-logical_group.
      ELSE.
        io_salv->get_metadata( ) .
      ENDIF.
    ELSE.
* If we are offline we need to build this.
      cl_salv_controller_metadata=>get_variant(
        EXPORTING
          r_layout  = lo_layout
        CHANGING
          s_variant = ls_vari ).
    ENDIF.

*... get the column information
    wt_fcat = cl_salv_controller_metadata=>get_lvc_fieldcatalog(
                           r_columns      = lo_columns
                           r_aggregations = lo_aggregations ).

*... get the layout information
    cl_salv_controller_metadata=>get_lvc_layout(
      EXPORTING
        r_functional_settings = lo_functional
        r_display_settings    = lo_display
        r_columns             = lo_columns
        r_aggregations        = lo_aggregations
      CHANGING
        s_layout              = ws_layo ).

* the fieldcatalog is not complete yet!
    CALL FUNCTION 'LVC_FIELDCAT_COMPLETE'
      EXPORTING
        i_complete       = 'X'
        i_refresh_buffer = space
        i_buffer_active  = space
        is_layout        = ws_layo
        i_test           = '1'
        i_fcat_complete  = 'X'
      IMPORTING
        es_layout        = ws_layo
      CHANGING
        ct_fieldcat      = wt_fcat.

    IF ls_vari IS NOT INITIAL AND
        ( io_salv->is_offline( ) = abap_true
          OR lv_intercept_data_active = abap_true ).
      CALL FUNCTION 'LVC_TRANSFER_TO_KKBLO'
        EXPORTING
          it_fieldcat_lvc         = wt_fcat
          is_layout_lvc           = ws_layo
        IMPORTING
          et_fieldcat_kkblo       = lt_kkblo_fieldcat
          es_layout_kkblo         = ls_kkblo_layout
        TABLES
          it_data                 = it_table
        EXCEPTIONS
          it_data_missing         = 1
          it_fieldcat_lvc_missing = 2
          OTHERS                  = 3.
      IF sy-subrc <> 0.
      ENDIF.

      CALL FUNCTION 'LT_VARIANT_LOAD'
        EXPORTING
          i_tabname           = '1'
          i_dialog            = ' '
          i_user_specific     = 'X'
          i_fcat_complete     = 'X'
        IMPORTING
          et_fieldcat         = lt_kkblo_fieldcat
          et_sort             = lt_kkblo_sort
          et_filter           = lt_kkblo_filter
        CHANGING
          cs_layout           = ls_kkblo_layout
          ct_default_fieldcat = lt_kkblo_fieldcat
          cs_variant          = ls_vari
        EXCEPTIONS
          wrong_input         = 1
          fc_not_complete     = 2
          not_found           = 3
          OTHERS              = 4.
      IF sy-subrc <> 0.
      ENDIF.

      CALL FUNCTION 'LVC_TRANSFER_FROM_KKBLO'
        EXPORTING
          it_fieldcat_kkblo = lt_kkblo_fieldcat
          it_sort_kkblo     = lt_kkblo_sort
          it_filter_kkblo   = lt_kkblo_filter
          is_layout_kkblo   = ls_kkblo_layout
        IMPORTING
          et_fieldcat_lvc   = wt_fcat
          et_sort_lvc       = wt_sort
          et_filter_lvc     = wt_filt
          es_layout_lvc     = ws_layo
        TABLES
          it_data           = it_table
        EXCEPTIONS
          it_data_missing   = 1
          OTHERS            = 2.
      IF sy-subrc <> 0.
      ENDIF.

    ELSE.
*  ... get the sort information
      wt_sort = cl_salv_controller_metadata=>get_lvc_sort( lo_sorts ).

*  ... get the filter information
      wt_filt = cl_salv_controller_metadata=>get_lvc_filter( lo_filters ).
    ENDIF.

  ENDMETHOD.


  METHOD zif_excel_converter~get_supported_class.
    rv_supported_class = 'CL_SALV_TABLE'.
  ENDMETHOD.

  METHOD zif_excel_converter~can_convert_object.

    DATA: lo_salv TYPE REF TO cl_salv_table.

    TRY.
        lo_salv ?= io_object.
      CATCH cx_sy_move_cast_error .
        RAISE EXCEPTION TYPE zcx_excel.
    ENDTRY.

  ENDMETHOD.


  METHOD zif_excel_converter~create_fieldcatalog.
    DATA: lo_salv TYPE REF TO cl_salv_table.

    zif_excel_converter~can_convert_object( io_object = io_object ).

    ws_option = is_option.

    lo_salv ?= io_object.

    CLEAR: es_layout,
           et_fieldcatalog,
           et_colors .

    IF lo_salv IS BOUND.
      load_data( EXPORTING io_salv   = lo_salv
                           it_table  = it_table ).
      apply_sort( EXPORTING it_table = it_table
                  IMPORTING eo_table = eo_table ) .

      get_color( EXPORTING io_table    = eo_table
                 IMPORTING et_colors   = et_colors ) .

      get_filter( IMPORTING et_filter  = et_filter
                  CHANGING  xo_table   = eo_table ) .

      update_catalog( CHANGING  cs_layout       = es_layout
                                ct_fieldcatalog = et_fieldcatalog ).
    ENDIF.
  ENDMETHOD.

  METHOD is_intercept_data_active.

    DATA: lr_s_type_runtime_info TYPE REF TO data.
    FIELD-SYMBOLS: <ls_type_runtime_info> TYPE any,
                   <lv_display>           TYPE any,
                   <lv_data>              TYPE any.

    rv_result = abap_false.
    TRY.
        CREATE DATA lr_s_type_runtime_info TYPE ('CL_SALV_BS_RUNTIME_INFO=>S_TYPE_RUNTIME_INFO').
        ASSIGN lr_s_type_runtime_info->* TO <ls_type_runtime_info>.
        CALL METHOD ('CL_SALV_BS_RUNTIME_INFO')=>('GET')
          RECEIVING
            value = <ls_type_runtime_info>.
        ASSIGN ('<LS_TYPE_RUNTIME_INFO>-DISPLAY') TO <lv_display>.
        CHECK sy-subrc = 0.
        ASSIGN ('<LS_TYPE_RUNTIME_INFO>-DATA') TO <lv_data>.
        CHECK sy-subrc = 0.
        IF <lv_display> = abap_false AND <lv_data> = abap_true.
          rv_result = abap_true.
        ENDIF.
      CATCH cx_sy_create_data_error cx_sy_dyn_call_error cx_salv_bs_sc_runtime_info.
        rv_result = abap_false.
    ENDTRY.

  ENDMETHOD.

ENDCLASS.
