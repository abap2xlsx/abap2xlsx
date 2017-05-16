class ZCL_EXCEL_WORKSHEET_PAGEBREAKS definition
  public
  create public .

public section.

  types:
    BEGIN OF ts_pagebreak_at ,
             cell_row  	 TYPE zexcel_cell_row,
             cell_column TYPE zexcel_cell_column,
           END OF ts_pagebreak_at .
  types:
    tt_pagebreak_at TYPE HASHED TABLE OF  ts_pagebreak_at WITH UNIQUE KEY cell_row cell_column .

  methods ADD_PAGEBREAK
    importing
      !IP_COLUMN type SIMPLE
      !IP_ROW type ZEXCEL_CELL_ROW
    raising
      ZCX_EXCEL .
  methods GET_ALL_PAGEBREAKS
    returning
      value(RT_PAGEBREAKS) type TT_PAGEBREAK_AT .
protected section.

  data MT_PAGEBREAKS type TT_PAGEBREAK_AT .
private section.
ENDCLASS.



CLASS ZCL_EXCEL_WORKSHEET_PAGEBREAKS IMPLEMENTATION.


METHOD add_pagebreak.
  DATA: ls_pagebreak      LIKE LINE OF me->mt_pagebreaks.

  ls_pagebreak-cell_row    = ip_row.
  ls_pagebreak-cell_column = zcl_excel_common=>convert_column2int( ip_column ).

  INSERT ls_pagebreak INTO TABLE me->mt_pagebreaks.


ENDMETHOD.


METHOD get_all_pagebreaks.
  rt_pagebreaks = me->mt_pagebreaks.
ENDMETHOD.
ENDCLASS.
