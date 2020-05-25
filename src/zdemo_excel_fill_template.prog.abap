*&---------------------------------------------------------------------*
*& Report  Fill Template
*&
*&---------------------------------------------------------------------*
*&
*&
*&---------------------------------------------------------------------*

REPORT zdemo_excel_fill_template.

TYPES
: BEGIN OF t_table1
,     person TYPE string
,     salary TYPE i
, END OF t_table1

, tt_table1 TYPE TABLE OF  t_table1 WITH EMPTY KEY

, BEGIN OF t_line1
,     carrid TYPE string
,     connid TYPE string
,     fldate TYPE string
,     price TYPE i
, END OF t_line1

, tt_line1 TYPE TABLE OF  t_line1 WITH EMPTY KEY

, BEGIN OF t_table2
,     carrid TYPE string
,     price TYPE i
,     line1 TYPE tt_line1
, END OF t_table2

, tt_table2 TYPE TABLE OF  t_table2 WITH EMPTY KEY

, BEGIN OF t_sheet1
,     date TYPE string
,     time TYPE string
,     user TYPE string
,     total TYPE i
,     price TYPE i
,     table1 TYPE tt_table1
,     table2 TYPE tt_table2
, END OF t_sheet1


, BEGIN OF t_table3
,     person TYPE string
,     salary TYPE i
, END OF t_table3

, tt_table3 TYPE TABLE OF  t_table3 WITH EMPTY KEY

, BEGIN OF t_sheet2
,     date TYPE string
,     time TYPE string
,     user TYPE string
,     total TYPE i
,     table3 TYPE tt_table3
, END OF t_sheet2
.



DATA
: lo_data TYPE REF TO zcl_excel_template_data
, gs_sheet1 TYPE   t_sheet1
, gs_sheet2 TYPE   t_sheet2
.

* define variables
DATA: lo_excel TYPE REF TO zcl_excel,
      reader   TYPE REF TO zif_excel_reader.


CONSTANTS: gc_save_file_name TYPE string VALUE 'fill_template_example.xlsx'.
INCLUDE zdemo_excel_outputopt_incl.

PARAMETERS: p_fpath TYPE string OBLIGATORY LOWER CASE DEFAULT 'C:\Users\sadfasdf\Desktop\abap2xlsx\ZABAP2XLSX_EXAMPLE.xlsx'.


PARAMETERS: p_file RADIOBUTTON GROUP rad1 DEFAULT 'X'
          , p_smw0 RADIOBUTTON GROUP rad1
          .

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_fpath.
  PERFORM get_file_path CHANGING p_fpath.


START-OF-SELECTION.

  CREATE OBJECT lo_data.

* generate data
  gs_sheet1 = VALUE #(
  date = |{ sy-datum DATE = ENVIRONMENT }|
  time = |{ sy-uzeit TIME = ENVIRONMENT }|
  user  = |{ sy-uname }|

  table1 = VALUE #(
                    ( person = 'Lurch Schpellchek' salary = '1200' )
                    ( person = 'Russell Sprout'    salary = '1300' )
                    ( person = 'Fergus Douchebag'  salary = '3000' )
                    ( person = 'Bartholomew Shoe'  salary = '100' )
                  )

  total = '5600'
  table2 = VALUE #(
                      ( line1 = VALUE #(
                                         (  carrid = 'AC' connid = '0820'  fldate = '20.12.2002' price = '1222'  )
                                       )
                        carrid ='AC'
                        price = '1222'
                      )
                      ( line1 = VALUE #(
                                          (  carrid = 'AF' connid = '0820'  fldate = '23.12.2002' price = '2222'  )
                                       )
                        carrid ='AF'
                        price = '2222'
                      )

                      ( line1 = VALUE #(
                                          (  carrid = 'LH' connid = '0400'  fldate = '28.02.1995' price = '899'  )
                                          (  carrid = 'LH' connid = '0454'  fldate = '17.11.1995' price = '1499'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '06.06.1995' price = '1090'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '28.04.1995' price = '6000'  )
                                          (  carrid = 'LH' connid = '9981'  fldate = '21.12.2002' price = '222'  )
                                       )
                        carrid ='LH'
                        price = '9488'
                      )
                      ( line1 = VALUE #(
                                          (  carrid = 'SQ' connid = '0026'  fldate = '28.02.1995' price = '849'  )
                                        )
                        carrid ='SQ'
                        price = '849'
                      )

                      ( line1 = VALUE #(
                                          (  carrid = 'LH' connid = '0400'  fldate = '28.02.1995' price = '899'  )
                                          (  carrid = 'LH' connid = '0454'  fldate = '17.11.1995' price = '1499'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '06.06.1995' price = '1090'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '28.04.1995' price = '6000'  )
                                          (  carrid = 'LH' connid = '9981'  fldate = '21.12.2002' price = '222'  )
                                       )
                        carrid ='LH'
                        price = '9488'
                      )
                      ( line1 = VALUE #(
                                          (  carrid = 'SQ' connid = '0026'  fldate = '28.02.1995' price = '849'  )
                                        )
                        carrid ='SQ'
                        price = '849'
                      )

                      ( line1 = VALUE #(
                                          (  carrid = 'LH' connid = '0400'  fldate = '28.02.1995' price = '899'  )
                                          (  carrid = 'LH' connid = '0454'  fldate = '17.11.1995' price = '1499'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '06.06.1995' price = '1090'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '28.04.1995' price = '6000'  )
                                          (  carrid = 'LH' connid = '9981'  fldate = '21.12.2002' price = '222'  )
                                       )
                        carrid ='LH'
                        price = '9488'
                      )
                      ( line1 = VALUE #(
                                          (  carrid = 'SQ' connid = '0026'  fldate = '28.02.1995' price = '849'  )
                                        )
                        carrid ='SQ'
                        price = '849'
                      )

                      ( line1 = VALUE #(
                                          (  carrid = 'LH' connid = '0400'  fldate = '28.02.1995' price = '899'  )
                                          (  carrid = 'LH' connid = '0454'  fldate = '17.11.1995' price = '1499'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '06.06.1995' price = '1090'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '28.04.1995' price = '6000'  )
                                          (  carrid = 'LH' connid = '9981'  fldate = '21.12.2002' price = '222'  )
                                       )
                        carrid ='LH'
                        price = '9488'
                      )
                      ( line1 = VALUE #(
                                          (  carrid = 'SQ' connid = '0026'  fldate = '28.02.1995' price = '849'  )
                                        )
                        carrid ='SQ'
                        price = '849'
                      )

                      ( line1 = VALUE #(
                                          (  carrid = 'LH' connid = '0400'  fldate = '28.02.1995' price = '899'  )
                                          (  carrid = 'LH' connid = '0454'  fldate = '17.11.1995' price = '1499'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '06.06.1995' price = '1090'  )
                                          (  carrid = 'LH' connid = '0455'  fldate = '28.04.1995' price = '6000'  )
                                          (  carrid = 'LH' connid = '9981'  fldate = '21.12.2002' price = '222'  )
                                       )
                        carrid ='LH'
                        price = '9488'
                      )
                      ( line1 = VALUE #(
                                          (  carrid = 'SQ' connid = '0026'  fldate = '28.02.1995' price = '849'  )
                                        )
                        carrid ='SQ'
                        price = '849'
                      )

  )

  price = '14003'
  ).


  gs_sheet2 = VALUE #(
  date = |{ sy-datum DATE = ENVIRONMENT }|
  time = |{ sy-uzeit TIME = ENVIRONMENT }|
  user  = |{ sy-uname }|

  table3 = VALUE #(
                    ( person = 'Lurch Schpellchek' salary = '1200' )
                    ( person = 'Russell Sprout'    salary = '1300' )
                    ( person = 'Fergus Douchebag'  salary = '3000' )
                    ( person = 'Bartholomew Shoe'  salary = '100' )
                  )

  total = '5600' ).

* add data
  lo_data->add( iv_sheet = 'Sheet1' iv_data = gs_sheet1 ).
  lo_data->add( iv_sheet = 'Sheet2' iv_data = gs_sheet2 ).

* create reader

  CREATE OBJECT reader TYPE zcl_excel_reader_2007.
* load template

  IF p_file IS NOT INITIAL.
    lo_excel = reader->load_file( p_fpath ).
  ELSE.
    lo_excel = reader->load_smw0( 'ZEXCEL_DEMO_TEMPLATE' ).
  ENDIF.


* merge data with template
  lo_excel->fill_template( lo_data ).

*** Create output
  lcl_output=>output( lo_excel ).


*&---------------------------------------------------------------------*
*&      Form  Get_file_path
*&---------------------------------------------------------------------*
FORM get_file_path CHANGING cv_path TYPE string.
  CLEAR cv_path.

  DATA:
    lv_rc          TYPE  i,
    lv_user_action TYPE  i,
    lt_file_table  TYPE  filetable,
    ls_file_table  LIKE LINE OF lt_file_table.

  cl_gui_frontend_services=>file_open_dialog(
  EXPORTING
    window_title        = 'select template  xlsx'
    multiselection      = ''
    default_extension   = '*.xlsx'
    file_filter         = 'Text file (*.xlsx)|*.xlsx|All (*.*)|*.*'
  CHANGING
    file_table          = lt_file_table
    rc                  = lv_rc
    user_action         = lv_user_action
  EXCEPTIONS
    OTHERS              = 1
    ).
  IF sy-subrc = 0.
    IF lv_user_action = cl_gui_frontend_services=>action_ok.
      IF lt_file_table IS NOT INITIAL.
        READ TABLE lt_file_table INTO ls_file_table INDEX 1.
        IF sy-subrc = 0.
          cv_path = ls_file_table-filename.
        ENDIF.
      ENDIF.
    ENDIF.
  ENDIF.
ENDFORM.                    " Get_file_path
