REPORT.

CONSTANTS: gc_save_file_name TYPE string VALUE 'ABAP2XLSX Inheritance.xlsx'.

*--------------------------------------------------------------------*
* Demo inheritance ZCL_EXCEL1
* Variation of ZCL_EXCEL that creates numerous sheets
*--------------------------------------------------------------------*
CLASS lcl_my_zcl_excel1 DEFINITION INHERITING FROM zcl_excel.
  PUBLIC SECTION.
    METHODS: constructor IMPORTING iv_sheetcount TYPE i DEFAULT 5.
ENDCLASS.

CLASS lcl_my_zcl_excel1 IMPLEMENTATION.
  METHOD constructor.
    DATA: lv_sheets_to_create TYPE i.
    super->constructor( ).
    lv_sheets_to_create = iv_sheetcount - 1. " one gets created by standard class
    DO lv_sheets_to_create TIMES.
      TRY.
          me->add_new_worksheet( ).
        CATCH zcx_excel.
      ENDTRY.
    ENDDO.
    me->set_active_sheet_index( 1 ).

  ENDMETHOD.
ENDCLASS.

*--------------------------------------------------------------------*
* Demo inheritance ZCL_EXCEL_WORKSHEET
* Variation of ZCL_EXCEL_WORKSHEET ( and ZCL_EXCEL that calls the new type of worksheet )
* that sets a fixed title
*--------------------------------------------------------------------*
CLASS lcl_my_zcl_excel2 DEFINITION INHERITING FROM zcl_excel.
  PUBLIC SECTION.
    METHODS: constructor.
ENDCLASS.

CLASS lcl_my_zcl_excel_worksheet DEFINITION INHERITING FROM zcl_excel_worksheet.
  PUBLIC SECTION.
    METHODS: constructor IMPORTING ip_excel TYPE REF TO zcl_excel
                                   ip_title TYPE zexcel_sheet_title OPTIONAL  " Will be ignored - keep parameter for demonstration purpose
                         RAISING   zcx_excel.
ENDCLASS.

CLASS lcl_my_zcl_excel2 IMPLEMENTATION.
  METHOD constructor.

    DATA: lo_worksheet TYPE REF TO zcl_excel_worksheet.

    super->constructor( ).

* To use own worksheet we have to remove the standard worksheet
    lo_worksheet = get_active_worksheet( ).
    me->worksheets->remove( lo_worksheet ).
* and replace it with own version
    CREATE OBJECT lo_worksheet TYPE lcl_my_zcl_excel_worksheet
      EXPORTING
        ip_excel = me
        ip_title = 'This title will be ignored'.
    me->worksheets->add( lo_worksheet ).

  ENDMETHOD.
ENDCLASS.

CLASS lcl_my_zcl_excel_worksheet IMPLEMENTATION.
  METHOD constructor.
    super->constructor( ip_excel = ip_excel
                        ip_title = 'Inherited Worksheet' ).

  ENDMETHOD.
ENDCLASS.

DATA: go_excel1 TYPE REF TO lcl_my_zcl_excel1.
DATA: go_excel2 TYPE REF TO lcl_my_zcl_excel2.


SELECTION-SCREEN BEGIN OF BLOCK bli WITH FRAME TITLE text-bli.
PARAMETERS: rbi_1 RADIOBUTTON GROUP rbi DEFAULT 'X' , " Simple inheritance
            rbi_2 RADIOBUTTON GROUP rbi.
SELECTION-SCREEN END OF BLOCK bli.

INCLUDE zdemo_excel_outputopt_incl.

END-OF-SELECTION.

  CASE 'X'.

    WHEN rbi_1.                                       " Simple inheritance of zcl_excel, object created directly
      CREATE OBJECT go_excel1
        EXPORTING
          iv_sheetcount = 5.
      lcl_output=>output( go_excel1 ).

    WHEN rbi_2.                                       " Inheritance of zcl_excel_worksheet, inheritance of zcl_excel needed to allow this
      CREATE OBJECT go_excel2.
      lcl_output=>output( go_excel2 ).


  ENDCASE.
