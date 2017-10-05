class ZCL_EXCEL_CONVERTER_SALV_TABLE definition
  public
  inheriting from ZCL_EXCEL_CONVERTER_ALV
  final
  create public .

*"* public components of class ZCL_EXCEL_CONVERTER_SALV_TABLE
*"* do not include other source files here!!!
public section.

  methods ZIF_EXCEL_CONVERTER~CAN_CONVERT_OBJECT
    redefinition .
  methods ZIF_EXCEL_CONVERTER~CREATE_FIELDCATALOG
    redefinition .
*"* protected components of class ZCL_EXCEL_CONVERTER_SALV_TABLE
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_CONVERTER_SALV_TABLE
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_CONVERTER_SALV_TABLE
*"* do not include other source files here!!!
protected section.
private section.

  methods LOAD_DATA
    importing
      !IO_SALV type ref to CL_SALV_TABLE
      !IT_TABLE type STANDARD TABLE .
ENDCLASS.



CLASS ZCL_EXCEL_CONVERTER_SALV_TABLE IMPLEMENTATION.


method LOAD_DATA.
  DATA: lo_columns        TYPE REF TO cl_salv_columns_table,
        lo_aggregations   TYPE REF TO cl_salv_aggregations,
        lo_sorts          TYPE REF TO cl_salv_sorts,
        lo_filters        TYPE REF TO cl_salv_filters,
        lo_functional     TYPE REF TO cl_salv_functional_settings,
        lo_display        TYPE REF TO cl_salv_display_settings.

  DATA: ls_vari   TYPE disvariant,
        lo_layout TYPE REF TO cl_salv_layout.

  DATA lt_kkblo_fieldcat TYPE kkblo_t_fieldcat.
  DATA ls_kkblo_layout   TYPE kkblo_layout.
  DATA lt_kkblo_filter   TYPE kkblo_t_filter.
  DATA lt_kkblo_sort     TYPE kkblo_t_sortinfo.

  lo_layout               = io_salv->get_layout( ) .
  lo_columns              = io_salv->get_columns( ).
  lo_aggregations         = io_salv->get_aggregations( ) .
  lo_sorts                = io_salv->get_sorts( ) .
  lo_filters              = io_salv->get_filters( ) .
  lo_display              = io_salv->get_display_settings( ) .
  lo_functional           = io_salv->get_functional_settings( ) .

  REFRESH: wt_fcat,
           wt_sort,
           wt_filt.

* First update metadata if we can.
  IF io_salv->is_offline( ) = abap_false.
    io_salv->get_metadata( ) .
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
     i_complete             = 'X'
     i_refresh_buffer       = space
     i_buffer_active        = space
     is_layout              = ws_layo
     i_test                 = '1'
     i_fcat_complete        = 'X'
   IMPORTING
*            E_EDIT                 =
     es_layout              = ws_layo
    CHANGING
      ct_fieldcat            = wt_fcat.

  IF ls_vari IS NOT INITIAL AND io_salv->is_offline( ) = abap_true.
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
* MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
    ENDIF.

    CALL FUNCTION 'LT_VARIANT_LOAD'
      EXPORTING
*       I_TOOL                      = 'LT'
        i_tabname                   = '1'
*       I_TABNAME_SLAVE             =
        i_dialog                    = ' '
        I_USER_SPECIFIC             = 'X'
*       I_DEFAULT                   = 'X'
*       I_NO_REPTEXT_OPTIMIZE       =
*       I_VIA_GRID                  =
        i_fcat_complete             = 'X'
      IMPORTING
*       E_EXIT                      =
        et_fieldcat                 = lt_kkblo_fieldcat
        et_sort                     = lt_kkblo_sort
        et_filter                   = lt_kkblo_filter
      CHANGING
        cs_layout                   = ls_kkblo_layout
        ct_default_fieldcat         = lt_kkblo_fieldcat
        cs_variant                  = ls_vari
     EXCEPTIONS
       wrong_input                 = 1
       fc_not_complete             = 2
       not_found                   = 3
       OTHERS                      = 4
              .
    IF sy-subrc <> 0.
* MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
    ENDIF.

    CALL FUNCTION 'LVC_TRANSFER_FROM_KKBLO'
      EXPORTING
*       I_TECH_COMPLETE                 =
*       I_STRUCTURE_NAME                =
        it_fieldcat_kkblo               = lt_kkblo_fieldcat
        it_sort_kkblo                   = lt_kkblo_sort
        it_filter_kkblo                 = lt_kkblo_filter
*       IT_SPECIAL_GROUPS_KKBLO         =
*       IT_FILTERED_ENTRIES_KKBLO       =
*       IT_GROUPLEVELS_KKBLO            =
*       IS_SUBTOT_OPTIONS_KKBLO         =
        is_layout_kkblo                 = ls_kkblo_layout
*       IS_REPREP_ID_KKBLO              =
*       I_CALLBACK_PROGRAM_KKBLO        =
*       IT_ADD_FIELDCAT                 =
*       IT_EXCLUDING_KKBLO              =
*       IT_EXCEPT_QINFO_KKBLO           =
      IMPORTING
        et_fieldcat_lvc                 = wt_fcat
        et_sort_lvc                     = wt_sort
        et_filter_lvc                   = wt_filt
*       ET_SPECIAL_GROUPS_LVC           =
*       ET_FILTER_INDEX_LVC             =
*       ET_GROUPLEVELS_LVC              =
*       ES_TOTAL_OPTIONS_LVC            =
        es_layout_lvc                   = ws_layo
*       ES_VARIANT_LVC                  =
*       E_VARIANT_SAVE_LVC              =
*       ES_PRINT_INFO_LVC               =
*       ES_REPREP_LVC                   =
*       E_REPREP_ACTIVE_LVC             =
*       ET_EXCLUDING_LVC                =
*       ET_EXCEPT_QINFO_LVC             =
      TABLES
        it_data                         = it_table
     EXCEPTIONS
       it_data_missing                 = 1
       OTHERS                          = 2
              .
    IF sy-subrc <> 0.
* MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*         WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
    ENDIF.

  ELSE.
*  ... get the sort information
    wt_sort = cl_salv_controller_metadata=>get_lvc_sort( lo_sorts ).

*  ... get the filter information
    wt_filt = cl_salv_controller_metadata=>get_lvc_filter( lo_filters ).
  ENDIF.

  endmethod.


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

  TRY.
    zif_excel_converter~can_convert_object( io_object = io_object ).
  ENDTRY.

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
ENDCLASS.
