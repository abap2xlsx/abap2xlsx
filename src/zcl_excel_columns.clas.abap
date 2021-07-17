class ZCL_EXCEL_COLUMNS definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_COLUMNS
*"* do not include other source files here!!!
  public section.
    types:
      begin of MTY_S_HASHED_COLUMN,
        COLUMN_INDEX type INT4,
        COLUMN       type ref to ZCL_EXCEL_COLUMN,
      end of MTY_S_HASHED_COLUMN ,
      MTY_TS_HASEHD_COLUMN type hashed table of MTY_S_HASHED_COLUMN with unique key COLUMN_INDEX.

    methods ADD
      importing
        !IO_COLUMN type ref to ZCL_EXCEL_COLUMN .
    methods CLEAR .
    methods CONSTRUCTOR .
    methods GET
      importing
        !IP_INDEX        type I
      returning
        value(EO_COLUMN) type ref to ZCL_EXCEL_COLUMN .
    methods GET_ITERATOR
      returning
        value(EO_ITERATOR) type ref to CL_OBJECT_COLLECTION_ITERATOR .
    methods IS_EMPTY
      returning
        value(IS_EMPTY) type FLAG .
    methods REMOVE
      importing
        !IO_COLUMN type ref to ZCL_EXCEL_COLUMN .
    methods SIZE
      returning
        value(EP_SIZE) type I .
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
  protected section.
*"* private components of class ZABAP_EXCEL_RANGES
*"* do not include other source files here!!!
  private section.

    data COLUMNS type ref to CL_OBJECT_COLLECTION .
    data COLUMNS_HASEHD type MTY_TS_HASEHD_COLUMN .
ENDCLASS.



CLASS ZCL_EXCEL_COLUMNS IMPLEMENTATION.


  method ADD.
    data: LS_HASHED_COLUMN type MTY_S_HASHED_COLUMN.

    LS_HASHED_COLUMN-COLUMN_INDEX = IO_COLUMN->GET_COLUMN_INDEX( ).
    LS_HASHED_COLUMN-COLUMN = IO_COLUMN.

    insert LS_HASHED_COLUMN into table COLUMNS_HASEHD .

    COLUMNS->ADD( IO_COLUMN ).
  endmethod.


  method CLEAR.
    clear COLUMNS_HASEHD.
    COLUMNS->CLEAR( ).
  endmethod.


  method CONSTRUCTOR.

    create object COLUMNS.

  endmethod.


  method GET.
    field-symbols: <LS_HASHED_COLUMN> type MTY_S_HASHED_COLUMN.

    read table COLUMNS_HASEHD with key COLUMN_INDEX = IP_INDEX assigning <LS_HASHED_COLUMN>.
    if SY-SUBRC = 0.
      EO_COLUMN = <LS_HASHED_COLUMN>-COLUMN.
    endif.
  endmethod.


  method GET_ITERATOR.
    EO_ITERATOR ?= COLUMNS->GET_ITERATOR( ).
  endmethod.


  method IS_EMPTY.
    IS_EMPTY = COLUMNS->IS_EMPTY( ).
  endmethod.


  method REMOVE.
    delete table COLUMNS_HASEHD with table key COLUMN_INDEX = IO_COLUMN->GET_COLUMN_INDEX( ) .
    COLUMNS->REMOVE( IO_COLUMN ).
  endmethod.


  method SIZE.
    EP_SIZE = COLUMNS->SIZE( ).
  endmethod.
ENDCLASS.
