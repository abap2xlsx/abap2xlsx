class ZCL_EXCEL_CONVERTER_RESULT_EX definition
  public
  inheriting from ZCL_EXCEL_CONVERTER_RESULT
  final
  create public .

*"* public components of class ZCL_EXCEL_CONVERTER_RESULT_EX
*"* do not include other source files here!!!
public section.

  methods ZIF_EXCEL_CONVERTER~CAN_CONVERT_OBJECT
    redefinition .
  methods ZIF_EXCEL_CONVERTER~CREATE_FIELDCATALOG
    redefinition .
*"* protected components of class ZCL_EXCEL_CONVERTER_RESULT_EX
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_CONVERTER_RESULT_EX
*"* do not include other source files here!!!
*"* protected components of class ZCL_EXCEL_CONVERTER_RESULT_EX
*"* do not include other source files here!!!
protected section.
*"* private components of class ZCL_EXCEL_CONVERTER_EX_RESULT
*"* do not include other source files here!!!
private section.
ENDCLASS.



CLASS ZCL_EXCEL_CONVERTER_RESULT_EX IMPLEMENTATION.


METHOD ZIF_EXCEL_CONVERTER~CAN_CONVERT_OBJECT.

  DATA: lo_result TYPE REF TO cl_salv_ex_result_data_table.

  TRY.
      lo_result ?= io_object.
    CATCH cx_sy_move_cast_error .
      RAISE EXCEPTION TYPE zcx_excel.
  ENDTRY.

ENDMETHOD.


METHOD zif_excel_converter~create_fieldcatalog.
  DATA: lo_result  TYPE REF TO cl_salv_ex_result_data_table,
        lo_ex_cm   TYPE REF TO cl_salv_ex_cm,
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

      lo_ex_cm ?= lo_result->r_model->r_model.
      ws_layo = lo_ex_cm->s_layo.
* T_DRDN  Instance Attribute  Public  Type  LVC_T_DROP
      wt_fcat = lo_ex_cm->t_fcat.
      wt_filt = lo_ex_cm->t_filt.
* T_HYPE  Instance Attribute  Public  Type  LVC_T_HYPE
* T_SELECTED_CELLS  Instance Attribute  Public  Type  LVC_T_CELL
* T_SELECTED_COLUMNS  Instance Attribute  Public  Type  LVC_T_COL
      wt_sort = lo_ex_cm->t_sort.

      apply_sort( EXPORTING it_table = <fs_table>
                  IMPORTING eo_table = eo_table ) .

      get_color( EXPORTING io_table    = eo_table
                 IMPORTING et_colors   = et_colors ) .

      get_filter( IMPORTING et_filter  = et_filter
                  CHANGING  xo_table   = eo_table ) .

      update_catalog( CHANGING  cs_layout       = es_layout
                                ct_fieldcatalog = et_fieldcatalog ).
    else.
* We have a problem and should stop here.
    ENDIF.
  ENDIF.
ENDMETHOD.
ENDCLASS.
