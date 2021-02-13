*----------------------------------------------------------------------*
*       CLASS ZCL_EXCEL_ROWS DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
class ZCL_EXCEL_ROWS definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_ROWS
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
  public section.
    types:
      begin of MTY_S_HASHED_ROW,
        ROW_INDEX type INT4,
        ROW       type ref to ZCL_EXCEL_ROW,
      end of MTY_S_HASHED_ROW ,
      MTY_TS_HASEHD_ROW type hashed table of MTY_S_HASHED_ROW with unique key ROW_INDEX.

    methods ADD
      importing
        !IO_ROW type ref to ZCL_EXCEL_ROW .
    methods CLEAR .
    methods CONSTRUCTOR .
    methods GET
      importing
        !IP_INDEX     type I
      returning
        value(EO_ROW) type ref to ZCL_EXCEL_ROW .
    methods GET_ITERATOR
      returning
        value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
    methods IS_EMPTY
      returning
        value(IS_EMPTY) type FLAG .
    methods REMOVE
      importing
        !IO_ROW type ref to ZCL_EXCEL_ROW .
    methods SIZE
      returning
        value(EP_SIZE) type I .
    methods GET_MIN_INDEX
      returning
        value(EP_INDEX) type I .
    methods GET_MAX_INDEX
      returning
        value(EP_INDEX) type I .
  protected section.
*"* private components of class ZABAP_EXCEL_RANGES
*"* do not include other source files here!!!
  private section.

    data ROWS type ref to CL_OBJECT_COLLECTION .
    data ROWS_HASEHD type MTY_TS_HASEHD_ROW .

ENDCLASS.



CLASS ZCL_EXCEL_ROWS IMPLEMENTATION.


  method ADD.
    data: LS_HASHED_ROW type MTY_S_HASHED_ROW.

    LS_HASHED_ROW-ROW_INDEX = IO_ROW->GET_ROW_INDEX( ).
    LS_HASHED_ROW-ROW = IO_ROW.

    insert LS_HASHED_ROW into table ROWS_HASEHD.

    ROWS->ADD( IO_ROW ).
  endmethod.                    "ADD


  method CLEAR.
    clear ROWS_HASEHD.
    ROWS->CLEAR( ).
  endmethod.                    "CLEAR


  method CONSTRUCTOR.

    create object ROWS.

  endmethod.                    "CONSTRUCTOR


  method GET.
    field-symbols: <LS_HASHED_ROW> type MTY_S_HASHED_ROW.

    read table ROWS_HASEHD with key ROW_INDEX = IP_INDEX assigning <LS_HASHED_ROW>.
    if SY-SUBRC = 0.
      EO_ROW = <LS_HASHED_ROW>-ROW.
    endif.
  endmethod.                    "GET


  method GET_ITERATOR.
    EO_ITERATOR ?= ROWS->GET_ITERATOR( ).
  endmethod.                    "GET_ITERATOR


  method IS_EMPTY.
    IS_EMPTY = ROWS->IS_EMPTY( ).
  endmethod.                    "IS_EMPTY


  method REMOVE.
    delete table ROWS_HASEHD with table key ROW_INDEX = IO_ROW->GET_ROW_INDEX( ) .
    ROWS->REMOVE( IO_ROW ).
  endmethod.                    "REMOVE


  method SIZE.
    EP_SIZE = ROWS->SIZE( ).
  endmethod.                    "SIZE

  METHOD get_min_index.
    FIELD-SYMBOLS: <ls_hashed_row> TYPE mty_s_hashed_row.

    LOOP AT rows_hasehd ASSIGNING <ls_hashed_row>.
      IF ep_index = 0 OR <ls_hashed_row>-row_index < ep_index.
        ep_index = <ls_hashed_row>-row_index.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.

  METHOD get_max_index.
    FIELD-SYMBOLS: <ls_hashed_row> TYPE mty_s_hashed_row.

    LOOP AT rows_hasehd ASSIGNING <ls_hashed_row>.
      IF <ls_hashed_row>-row_index > ep_index.
        ep_index = <ls_hashed_row>-row_index.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.

ENDCLASS.
